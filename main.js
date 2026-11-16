// To Do Panel: a sidebar listing Microsoft To Do tasks, read through the
// second-brain CLI so the plugin reuses its sign-in and never holds a token.
// Dragging a task into a note inserts it as a Markdown line; clicking a task
// inserts it at the cursor of the last active note.
//
// Plain JavaScript with no build step: copy this folder into the vault's
// plugins folder (see README.md).

const {
  ItemView,
  MarkdownView,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  setIcon,
} = require("obsidian");
const { execFile } = require("child_process");

const VIEW_TYPE = "second-brain-todo-panel";
const DRAG_TYPE = "application/x-second-brain-todo";

const OLD_DEFAULT_COMMAND = 'wsl.exe -e bash -lc "second-brain todo-lists --json"';

const DEFAULT_SETTINGS = {
  // Run inside WSL through a login shell, so ~/.local/bin is on PATH.
  command: 'wsl.exe -e bash -lc "second-brain todo-lists --json --completed-days 7"',
  template: "- [ ] {{title}} {{due}}",
  dueFormat: "📅 {{date}}",
  linkFormat: "https://to-do.office.com/tasks/id/{{id}}/details",
  hideEmptyLists: true,
  timeoutSeconds: 60,
  // Panel state, kept between sessions.
  tab: "lists",
  showCompleted: false,
  collapsed: { lists: [], dates: ["date:none"] },
};

module.exports = class TodoPanelPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    // Saved settings from 0.1 carry the old default command, which left out
    // completed tasks; move them onto the new default.
    if (this.settings.command === OLD_DEFAULT_COMMAND) this.settings.command = DEFAULT_SETTINGS.command;
    this.registerView(VIEW_TYPE, (leaf) => new TodoPanelView(leaf, this));
    this.addRibbonIcon("list-checks", "Open To Do panel", () => this.openPanel());
    this.addCommand({
      id: "open-panel",
      name: "Open panel",
      callback: () => this.openPanel(),
    });
    this.addCommand({
      id: "refresh",
      name: "Refresh tasks",
      callback: () => this.panels().forEach((view) => view.refresh()),
    });
    this.addSettingTab(new TodoPanelSettingTab(this.app, this));

    // Keep the "in note" markers in step with the note being worked on.
    const noteChanged = () => this.panels().forEach((view) => view.updateNoteMatches());
    this.registerEvent(this.app.workspace.on("active-leaf-change", noteChanged));
    this.registerEvent(this.app.workspace.on("file-open", noteChanged));
    let pending = null;
    this.registerEvent(
      this.app.workspace.on("editor-change", () => {
        window.clearTimeout(pending);
        pending = window.setTimeout(noteChanged, 300);
      })
    );
    this.registerEvent(
      this.app.vault.on("modify", (file) => {
        if (file === this.app.workspace.getActiveFile()) noteChanged();
      })
    );
  }

  onunload() {
    // Obsidian detaches the views itself; nothing else is held open.
  }

  panels() {
    return this.app.workspace
      .getLeavesOfType(VIEW_TYPE)
      .map((leaf) => leaf.view)
      .filter((view) => view instanceof TodoPanelView);
  }

  async openPanel() {
    let leaf = this.app.workspace.getLeavesOfType(VIEW_TYPE)[0];
    if (!leaf) {
      leaf = this.app.workspace.getRightLeaf(false);
      await leaf.setViewState({ type: VIEW_TYPE, active: true });
    }
    this.app.workspace.revealLeaf(leaf);
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  // Runs the configured command and resolves to the parsed JSON envelope.
  fetchLists() {
    const [file, ...args] = splitCommand(this.settings.command);
    if (!file) return Promise.reject(new Error("No command configured."));
    return new Promise((resolve, reject) => {
      execFile(
        file,
        args,
        {
          timeout: Math.max(5, Number(this.settings.timeoutSeconds) || 60) * 1000,
          windowsHide: true,
          maxBuffer: 32 * 1024 * 1024,
        },
        (error, stdout, stderr) => {
          let payload = null;
          try {
            payload = JSON.parse(stdout);
          } catch (_) {
            // Fall through: report the process error or its stderr.
          }
          if (payload && payload.error) return reject(new Error(payload.error));
          if (payload && Array.isArray(payload.data)) return resolve(payload);
          const detail = lastLines(stderr) || (error && error.message) || "No output.";
          reject(new Error(`Could not read tasks.\n${detail}`));
        }
      );
    });
  }

  // The Markdown a task becomes when dropped or inserted.
  renderTask(task, list) {
    const s = this.settings;
    const values = {
      title: task.title,
      list: list.name,
      id: task.id,
      date: task.due_date || "",
      due: task.due_date ? fill(s.dueFormat, { date: task.due_date }) : "",
      link: fill(s.linkFormat, { id: encodeURIComponent(task.id), list: encodeURIComponent(list.id) }),
    };
    let line = fill(s.template, values);
    if (task.completed) line = line.replace("[ ]", "[x]");
    return line
      .replace(/[ \t]{2,}/g, " ")
      .replace(/[ \t]+$/gm, "");
  }

  // The date a daily note stands for, as YYYY-MM-DD, or null for any other
  // note. Uses the Daily notes core plugin's date format (YYYY-MM-DD unless
  // set otherwise) on the file name.
  dailyNoteDate(file) {
    if (!file || !file.basename || !window.moment) return null;
    let format = "YYYY-MM-DD";
    try {
      const daily = this.app.internalPlugins.getPluginById("daily-notes");
      const configured = daily && daily.instance && daily.instance.options && daily.instance.options.format;
      if (configured) format = configured;
    } catch (_) {
      // Internal API; fall back to the default format.
    }
    // A format with folders ("YYYY/MM/YYYY-MM-DD") names the path, so match
    // its last part against the file name.
    const nameFormat = format.split("/").pop();
    const parsed = window.moment(file.basename, nameFormat, true);
    return parsed.isValid() ? parsed.format("YYYY-MM-DD") : null;
  }

  // The note the user is working in: the active Markdown editor, which stays
  // set while focus is in the sidebar. Resolves to { file, text } or null.
  async activeNote() {
    const active = this.app.workspace.activeEditor;
    if (active && active.editor) return { file: active.file, text: active.editor.getValue() };
    const file = this.app.workspace.getActiveFile();
    if (file && file.extension === "md") return { file, text: await this.app.vault.cachedRead(file) };
    return null;
  }

  // Click-to-insert: put the task on its own line at the cursor of the note
  // that was active last. activeEditor survives focus moving to the sidebar.
  insertIntoActiveNote(text) {
    const editor =
      this.app.workspace.activeEditor?.editor ||
      this.app.workspace.getActiveViewOfType(MarkdownView)?.editor;
    if (!editor) {
      new Notice("Click into a note first, then click the task.");
      return;
    }
    const cursor = editor.getCursor();
    const line = editor.getLine(cursor.line);
    if (line.trim() === "") {
      editor.replaceRange(text, { line: cursor.line, ch: 0 }, { line: cursor.line, ch: line.length });
      editor.setCursor({ line: cursor.line, ch: text.length });
    } else {
      editor.replaceRange("\n" + text, { line: cursor.line, ch: line.length });
      editor.setCursor({ line: cursor.line + 1, ch: text.length });
    }
    editor.focus();
  }
};

class TodoPanelView extends ItemView {
  constructor(leaf, plugin) {
    super(leaf);
    this.plugin = plugin;
    this.lists = [];
    this.filter = "";
    this.loading = false;
    this.error = null;
    this.fetchedAt = null;
  }

  getViewType() {
    return VIEW_TYPE;
  }

  getDisplayText() {
    return "To Do";
  }

  getIcon() {
    return "list-checks";
  }

  async onOpen() {
    const root = this.contentEl;
    root.empty();
    root.addClass("sb-todo-panel");

    const toolbar = root.createDiv({ cls: "sb-todo-toolbar" });
    const search = toolbar.createEl("input", {
      type: "search",
      attr: { placeholder: "Filter tasks…" },
    });
    search.addEventListener("input", () => {
      // A filter opens every group with a match, using its own collapsed
      // state, so clearing it brings the groups back as they were.
      this.filter = search.value.trim().toLowerCase();
      this.filterCollapsed = new Set();
      this.renderLists();
    });
    this.completedEl = toolbar.createEl("button", { cls: "sb-todo-toggle" });
    this.completedEl.addEventListener("click", () => {
      this.plugin.settings.showCompleted = !this.plugin.settings.showCompleted;
      this.plugin.saveSettings();
      this.renderLists();
    });
    this.toggleAllEl = toolbar.createEl("button");
    this.toggleAllEl.addEventListener("click", () => this.toggleAll());
    const refresh = toolbar.createEl("button", { attr: { "aria-label": "Refresh" } });
    setIcon(refresh, "refresh-cw");
    refresh.addEventListener("click", () => this.refresh());

    const tabs = root.createDiv({ cls: "sb-todo-tabs", attr: { role: "tablist" } });
    this.tabEls = {};
    for (const [key, label] of [["lists", "Lists"], ["dates", "By date"]]) {
      const tab = tabs.createEl("button", { cls: "sb-todo-tab", text: label, attr: { role: "tab" } });
      tab.addEventListener("click", () => this.setTab(key));
      this.tabEls[key] = tab;
    }

    this.statusEl = root.createDiv({ cls: "sb-todo-status" });
    this.listsEl = root.createDiv();
    this.renderLists();
    await this.refresh();
  }

  async refresh() {
    if (this.loading) return;
    this.loading = true;
    this.error = null;
    this.renderStatus();
    try {
      const payload = await this.plugin.fetchLists();
      this.lists = payload.data;
      this.fetchedAt = new Date();
    } catch (e) {
      this.error = e.message;
    } finally {
      this.loading = false;
      this.renderLists();
    }
  }

  renderStatus() {
    if (!this.statusEl) return;
    this.statusEl.empty();
    if (this.loading) {
      this.statusEl.setText("Loading tasks…");
    } else if (this.error) {
      this.statusEl.createDiv({ cls: "sb-todo-error", text: this.error });
    } else if (this.fetchedAt) {
      const time = this.fetchedAt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
      this.statusEl.setText(`Updated ${time} · drag a task into a note, or click to insert`);
    }
  }

  get tab() {
    return this.plugin.settings.tab === "dates" ? "dates" : "lists";
  }

  setTab(tab) {
    this.plugin.settings.tab = tab;
    this.plugin.saveSettings();
    this.renderLists();
  }

  // Collapsed group keys for the current tab; a filter keeps its own set.
  collapsedSet() {
    if (this.filter) return this.filterCollapsed || (this.filterCollapsed = new Set());
    const all = this.plugin.settings.collapsed || (this.plugin.settings.collapsed = {});
    return new Set(all[this.tab] || []);
  }

  setCollapsed(set) {
    if (this.filter) {
      this.filterCollapsed = set;
      return;
    }
    this.plugin.settings.collapsed = { ...this.plugin.settings.collapsed, [this.tab]: [...set] };
    this.plugin.saveSettings();
  }

  toggleAll() {
    const groups = this.groups();
    const collapsed = this.collapsedSet();
    const anyOpen = groups.some((g) => !collapsed.has(g.key));
    this.setCollapsed(anyOpen ? new Set(groups.map((g) => g.key)) : new Set());
    this.renderLists();
  }

  // The groups the current tab shows, each with its matching tasks in order.
  groups() {
    const showCompleted = this.plugin.settings.showCompleted;
    // Open tasks first in every group; completed ones, when shown, follow.
    const openFirst = (items) => [
      ...items.filter((e) => !e.task.completed),
      ...items.filter((e) => e.task.completed),
    ];
    const entries = (list) =>
      list.tasks
        .filter((task) => (showCompleted || !task.completed) && this.matches(task, list))
        .map((task) => ({ task, list }));

    if (this.tab === "lists") {
      return this.lists
        .map((list) => ({ key: list.id, name: list.name, items: openFirst(entries(list)) }))
        .filter((g) => g.items.length > 0 || (!this.filter && !this.plugin.settings.hideEmptyLists));
    }

    // The day the sections are relative to: the active daily note's date, or
    // today when the active note is not a daily note.
    const today = this.refDate || localToday();
    const all = this.lists.flatMap(entries);
    const byTitle = (a, b) => a.task.title.localeCompare(b.task.title);
    const byDue = (a, b) => a.task.due_date.localeCompare(b.task.due_date) || byTitle(a, b);
    const important = (a, b) => (b.task.importance === "high") - (a.task.importance === "high");
    const groups = [
      {
        key: "date:today",
        name: today === localToday() ? "Today" : longDate(today),
        items: all.filter((e) => e.task.due_date === today).sort((a, b) => important(a, b) || byTitle(a, b)),
      },
      { key: "date:later", name: "Later", items: all.filter((e) => e.task.due_date > today).sort(byDue) },
      { key: "date:earlier", name: "Earlier", items: all.filter((e) => e.task.due_date && e.task.due_date < today).sort(byDue) },
      { key: "date:none", name: "No date", items: all.filter((e) => !e.task.due_date).sort(byTitle) },
    ];
    for (const g of groups) g.items = openFirst(g.items);
    return groups.filter((g) => g.items.length > 0 || (!this.filter && g.key !== "date:none"));
  }

  renderLists() {
    this.renderStatus();
    if (!this.listsEl) return;
    this.listsEl.empty();
    if (this.completedEl) {
      const on = this.plugin.settings.showCompleted;
      this.completedEl.toggleClass("is-active", on);
      this.completedEl.setAttr("aria-label", on ? "Hide completed tasks" : "Show completed tasks (last 7 days)");
      this.completedEl.setAttr("aria-pressed", String(on));
      setIcon(this.completedEl, "circle-check");
    }
    for (const [key, el] of Object.entries(this.tabEls || {})) {
      el.toggleClass("is-active", key === this.tab);
      el.setAttr("aria-selected", String(key === this.tab));
    }

    const today = localToday();
    const groups = this.groups();
    const collapsed = this.collapsedSet();

    for (const group of groups) {
      const details = this.listsEl.createEl("details", { cls: "sb-todo-list" });
      details.open = !collapsed.has(group.key);
      details.addEventListener("toggle", () => {
        const set = this.collapsedSet();
        if (details.open) set.delete(group.key);
        else set.add(group.key);
        this.setCollapsed(set);
        this.updateToggleAll();
      });
      const summary = details.createEl("summary", { text: group.name });
      const done = group.items.filter((e) => e.task.completed).length;
      summary.createSpan({ cls: "sb-todo-count", text: String(group.items.length - done) });
      if (done) summary.createSpan({ cls: "sb-todo-done-count", text: `+${done} done` });
      if (group.items.length === 0) {
        details.createDiv({ cls: "sb-todo-empty", text: "Nothing here" });
      }
      for (const { task, list } of group.items) {
        this.renderTask(details, task, list, today, this.tab === "dates");
      }
    }
    this.updateToggleAll();
    this.updateNoteMatches();
  }

  updateToggleAll() {
    if (!this.toggleAllEl) return;
    const collapsed = this.collapsedSet();
    const anyOpen = this.groups().some((g) => !collapsed.has(g.key));
    this.toggleAllEl.setAttr("aria-label", anyOpen ? "Collapse all" : "Expand all");
    setIcon(this.toggleAllEl, anyOpen ? "chevrons-down-up" : "chevrons-up-down");
  }

  renderTask(parent, task, list, today, showList = false) {
    const row = parent.createDiv({ cls: "sb-todo-task", attr: { draggable: "true" } });
    // A marker per task, with wrapped title lines hanging beside it, so a
    // continuation line never looks like the start of the next task.
    const marker = row.createSpan({ cls: "sb-todo-marker", attr: { "aria-hidden": "true" } });
    if (task.importance === "high") marker.addClass("is-important");
    if (task.completed) row.addClass("is-completed");

    const body = row.createDiv({ cls: "sb-todo-body" });
    body.createDiv({ cls: "sb-todo-title", text: task.title });

    const meta = [];
    if (task.completed) meta.push(["sb-todo-done", task.completed_date ? `Done ${shortDate(task.completed_date)}` : "Done"]);
    else if (task.importance === "high") meta.push(["sb-todo-flag", "Important"]);
    if (task.due_date && task.completed) {
      meta.push(["sb-todo-due", `Due ${shortDate(task.due_date)}`]);
    } else if (task.due_date) {
      const state = task.due_date < today ? " is-overdue" : task.due_date === today ? " is-today" : "";
      meta.push([`sb-todo-due${state}`, shortDate(task.due_date)]);
    }
    if (showList) meta.push(["sb-todo-list-name", list.name]);
    if (task.notes) meta.push(["sb-todo-has-notes", "Notes"]);
    const metaEl = body.createDiv({ cls: "sb-todo-meta" });
    // Shown by CSS only while the active note already has this task.
    metaEl.createSpan({ cls: "sb-todo-in-note", text: "In note" });
    for (const [cls, text] of meta) metaEl.createSpan({ cls, text });
    row.sbEntry = { task, list };
    this.markNoteMatch(row);
    if (task.notes) row.setAttr("title", task.notes);

    row.addEventListener("dragstart", (evt) => {
      const markdown = this.plugin.renderTask(task, list);
      // The editor inserts text/plain at the drop point on its own.
      evt.dataTransfer.setData("text/plain", markdown);
      evt.dataTransfer.setData(DRAG_TYPE, JSON.stringify({ id: task.id, list_id: list.id }));
      row.addClass("is-dragging");
    });
    row.addEventListener("dragend", () => row.removeClass("is-dragging"));
    row.addEventListener("click", () => {
      this.plugin.insertIntoActiveNote(this.plugin.renderTask(task, list));
    });
  }

  // Re-reads the active note and marks the tasks it already contains,
  // without rebuilding the panel, so scrolling and open groups stay put.
  async updateNoteMatches() {
    const note = await this.plugin.activeNote();
    this.noteLines = note ? noteTaskLines(note.text) : null;
    this.noteName = note && note.file ? note.file.basename : "";
    const refDate = (note && this.plugin.dailyNoteDate(note.file)) || null;
    if (refDate !== this.refDate) {
      this.refDate = refDate;
      // Sections depend on the day; rebuilding calls back here once more,
      // when the day no longer differs.
      if (this.tab === "dates") return this.renderLists();
    }
    if (!this.listsEl) return;
    for (const row of this.listsEl.querySelectorAll(".sb-todo-task")) this.markNoteMatch(row);
  }

  markNoteMatch(row) {
    if (!row.sbEntry) return;
    const state = this.inNote(row.sbEntry.task, row.sbEntry.list);
    row.toggleClass("is-in-note", Boolean(state));
    row.toggleClass("is-checked-in-note", state === "checked");
    const badge = row.querySelector(".sb-todo-in-note");
    if (!badge) return;
    const where = this.noteName || "this note";
    badge.setAttr("title", state === "checked" ? `Checked off in ${where}` : state ? `Already in ${where}` : "");
  }

  // Whether the active note holds this task: null when it does not, else
  // "checked" when a matching line is ticked with [x], or "open". A line
  // matches when, list marker and checkbox aside, it is the task's title (see
  // normalizeText), or exactly what dropping the task would insert.
  inNote(task, list) {
    const lines = this.noteLines;
    if (!lines) return null;
    const title = normalizeText(task.title);
    if (title && lines.titles.has(title)) return lines.titles.get(title) ? "checked" : "open";
    const rendered = this.plugin.renderTask(task, list).trim();
    return lines.raw.has(rendered) ? "open" : null;
  }

  matches(task, list) {
    if (!this.filter) return true;
    const haystack = `${task.title} ${task.notes || ""} ${list.name}`.toLowerCase();
    return this.filter.split(/\s+/).every((word) => haystack.includes(word));
  }
}

class TodoPanelSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    const s = this.plugin.settings;
    containerEl.empty();

    this.text(
      "Command",
      "Runs the second-brain CLI and must print `todo-lists --json` output. " +
        "The default runs it inside WSL.",
      "command"
    );
    this.text(
      "Task template",
      "What a dropped task becomes. Placeholders: {{title}}, {{due}}, {{date}}, {{list}}, {{link}}, {{id}}.",
      "template"
    );
    this.text("Due format", "What {{due}} becomes when the task has a due date. {{date}} is YYYY-MM-DD.", "dueFormat");
    this.text("Link format", "What {{link}} becomes. Placeholders: {{id}}, {{list}}.", "linkFormat");

    new Setting(containerEl)
      .setName("Hide empty lists")
      .addToggle((toggle) =>
        toggle.setValue(s.hideEmptyLists).onChange(async (value) => {
          s.hideEmptyLists = value;
          await this.plugin.saveSettings();
          this.plugin.panels().forEach((view) => view.renderLists());
        })
      );

    new Setting(containerEl)
      .setName("Timeout (seconds)")
      .setDesc("How long to wait for the command before giving up.")
      .addText((text) =>
        text.setValue(String(s.timeoutSeconds)).onChange(async (value) => {
          const seconds = parseInt(value, 10);
          if (Number.isFinite(seconds) && seconds > 0) {
            s.timeoutSeconds = seconds;
            await this.plugin.saveSettings();
          }
        })
      );
  }

  text(name, desc, key) {
    new Setting(this.containerEl)
      .setName(name)
      .setDesc(desc)
      .addText((text) =>
        text.setValue(this.plugin.settings[key]).onChange(async (value) => {
          this.plugin.settings[key] = value;
          await this.plugin.saveSettings();
        })
      );
  }
}

// Splits a command line into program and arguments, honouring double and
// single quotes, so arguments go to execFile without a Windows shell.
function splitCommand(command) {
  const parts = [];
  const re = /"([^"]*)"|'([^']*)'|(\S+)/g;
  let match;
  while ((match = re.exec(command || "")) !== null) {
    parts.push(match[1] ?? match[2] ?? match[3]);
  }
  return parts;
}

// The task-like lines of a note: each raw trimmed line, and each list item's
// text reduced by normalizeText, mapped to whether any such line is checked.
// Only [x] or [X] counts as checked; [<], [>], [-] and other marks do not.
function noteTaskLines(text) {
  const raw = new Set();
  const titles = new Map();
  for (const line of (text || "").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    raw.add(trimmed);
    const item = trimmed.match(/^(?:[-*+]|\d+[.)])\s+(?:\[(.)\]\s+)?(.*)$/);
    if (!item) continue;
    const key = normalizeText(item[2]);
    const checked = item[1] === "x" || item[1] === "X";
    titles.set(key, Boolean(titles.get(key)) || checked);
  }
  return { raw, titles };
}

// What makes two task texts "the same": case, indentation and runs of
// whitespace do not count, and neither does a tail of emojis, dates and links,
// such as "📅 2026-11-20", "✅ 16/11/2026" or "[To Do](https://…)".
const TRAILING_NOISE = new RegExp(
  "(?:\\s*(?:" +
    "\\[[^\\]]*\\]\\([^)]*\\)" + // Markdown link
    "|\\d{4}-\\d{2}-\\d{2}(?:[ T]\\d{1,2}:\\d{2})?" + // 2026-11-20, optional time
    "|\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}" + // 20/11/2026, 20.11.26
    "|[\\p{Extended_Pictographic}\\u2190-\\u21FF\\u2B00-\\u2BFF\\uFE0F\\u200D]+" + // emojis and arrows
    "))+\\s*$",
  "u"
);

function normalizeText(text) {
  return (text || "").replace(TRAILING_NOISE, "").replace(/\s+/g, " ").trim().toLowerCase();
}

function fill(template, values) {
  return (template || "").replace(/\{\{\s*(\w+)\s*\}\}/g, (whole, key) =>
    Object.prototype.hasOwnProperty.call(values, key) ? values[key] : whole
  );
}

function lastLines(text, count = 5) {
  return (text || "").trim().split(/\r?\n/).slice(-count).join("\n");
}

function localToday() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

function longDate(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  const options = { weekday: "short", day: "numeric", month: "short" };
  if (year !== new Date().getFullYear()) options.year = "numeric";
  return date.toLocaleDateString([], options);
}

function shortDate(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  const sameYear = year === new Date().getFullYear();
  return date.toLocaleDateString([], sameYear ? { day: "numeric", month: "short" } : { day: "numeric", month: "short", year: "numeric" });
}

