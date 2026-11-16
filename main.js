// To Do Panel: a sidebar listing Microsoft To Do tasks. The plugin signs in
// to Microsoft itself (OAuth device code flow, read-only Tasks permission) and
// reads tasks from Microsoft Graph. Dragging a task into a note inserts it as
// a Markdown line; clicking a task inserts it at the cursor of the last
// active note.
//
// Plain JavaScript with no build step: copy manifest.json, main.js and
// styles.css into the vault's plugins folder (see README.md).

const {
  ItemView,
  MarkdownView,
  Modal,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  requestUrl,
  setIcon,
} = require("obsidian");

const VIEW_TYPE = "second-brain-todo-panel";
const DRAG_TYPE = "application/x-second-brain-todo";
const GRAPH = "https://graph.microsoft.com/v1.0";
// Where the sign-in session lives: Obsidian's secret storage, kept on this
// device and outside the vault, so it is never synced or committed.
const SESSION_SECRET_ID = "todo-panel-microsoft-session";

const DEFAULT_SETTINGS = {
  // Microsoft sign-in. There is no shared app registration: each user
  // registers their own (see README.md) and pastes its IDs here.
  clientId: "",
  tenant: "common",
  scopes: "Tasks.Read",
  completedDays: 7,
  template: "- [ ] {{title}} {{due}}",
  dueFormat: "📅 {{date}}",
  linkFormat: "https://to-do.office.com/tasks/id/{{id}}/details",
  hideEmptyLists: true,
  // Panel state, kept between sessions.
  tab: "lists",
  showCompleted: false,
  collapsed: { lists: [], dates: ["date:none"] },
};

// Settings from the versions that read tasks through an external command.
const RETIRED_SETTINGS = ["command", "timeoutSeconds"];

module.exports = class TodoPanelPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    for (const key of RETIRED_SETTINGS) delete this.settings[key];
    this.auth = new MicrosoftAuth(this);
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
    // Stops a sign-in that is still waiting for the user.
    if (this.signInModal) this.signInModal.close();
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

  // Reads every To Do list with its open tasks (and recently completed ones)
  // from Microsoft Graph.
  fetchLists() {
    return fetchTodoLists(this.auth, Math.max(0, Number(this.settings.completedDays) || 0));
  }

  // Opens the sign-in dialog; it asks for the app's IDs first when missing.
  signIn() {
    if (this.signInModal) return;
    this.signInModal = new SignInModal(this.app, this);
    this.signInModal.open();
  }

  signOut() {
    this.auth.signOut();
    this.panels().forEach((view) => view.clear());
    new Notice("Signed out of Microsoft.");
  }

  // Called by the sign-in dialog once Microsoft has issued tokens.
  signedIn() {
    const who = this.auth.account ? ` as ${this.auth.account}` : "";
    new Notice(`Signed in to Microsoft${who}.`);
    this.panels().forEach((view) => view.refresh());
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
    if (!this.plugin.auth.signedIn) {
      this.clear();
      return;
    }
    this.loading = true;
    this.error = null;
    this.renderStatus();
    try {
      this.lists = await this.plugin.fetchLists();
      this.fetchedAt = new Date();
    } catch (e) {
      this.error = e.message;
      if (!this.plugin.auth.signedIn) this.lists = [];
    } finally {
      this.loading = false;
      this.renderLists();
    }
  }

  // Back to the signed-out state: no tasks, and a way to sign in.
  clear() {
    this.lists = [];
    this.fetchedAt = null;
    this.error = null;
    this.renderLists();
  }

  renderStatus() {
    if (!this.statusEl) return;
    this.statusEl.empty();
    if (!this.plugin.auth.signedIn && !this.loading) {
      const box = this.statusEl.createDiv({ cls: "sb-todo-signin" });
      if (this.error) box.createDiv({ cls: "sb-todo-error", text: this.error });
      box.createDiv({ text: "Sign in to Microsoft to show your To Do tasks." });
      const button = box.createEl("button", { cls: "mod-cta", text: "Sign in" });
      button.addEventListener("click", () => this.plugin.signIn());
    } else if (this.loading) {
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
    const auth = this.plugin.auth;
    containerEl.empty();

    new Setting(containerEl).setName("Microsoft account").setHeading();
    const account = new Setting(containerEl);
    if (auth.signedIn) {
      account
        .setName(auth.account ? `Signed in as ${auth.account}` : "Signed in")
        .setDesc("The sign-in is kept in Obsidian's secret storage on this device, never in the vault.")
        .addButton((button) =>
          button.setButtonText("Sign out").onClick(() => {
            this.plugin.signOut();
            this.display();
          })
        );
    } else {
      account
        .setName("Not signed in")
        .setDesc("Add your app's client ID below, then sign in with a code in your browser.")
        .addButton((button) =>
          button
            .setButtonText("Sign in")
            .setCta()
            .onClick(() => this.plugin.signIn())
        );
    }

    // Changing which app or directory to sign in with ends the current
    // sign-in, which belongs to the old one.
    const authSetting = (name, desc, key, placeholder) =>
      new Setting(containerEl)
        .setName(name)
        .setDesc(desc)
        .addText((text) =>
          text
            .setPlaceholder(placeholder)
            .setValue(s[key])
            .onChange(async (value) => {
              s[key] = value.trim();
              await this.plugin.saveSettings();
              if (auth.signedIn && !auth.matchesSettings()) {
                this.plugin.signOut();
                this.display();
              }
            })
        );
    authSetting(
      "Client ID",
      "The Application (client) ID of the app you registered in Microsoft Entra. See the README.",
      "clientId",
      "00000000-0000-0000-0000-000000000000"
    );
    authSetting(
      "Tenant",
      "Your directory (tenant) ID if the app accepts only accounts in your organisation. " +
        "Otherwise common, organizations or consumers, matching the app's supported account types.",
      "tenant",
      "common"
    );
    authSetting(
      "Permissions",
      "Microsoft Graph permissions to ask for. Tasks.Read is enough; the plugin never changes tasks.",
      "scopes",
      "Tasks.Read"
    );

    new Setting(containerEl).setName("Tasks").setHeading();
    new Setting(containerEl)
      .setName("Completed tasks to load (days)")
      .setDesc("How far back the show-completed toggle reaches. 0 loads no completed tasks.")
      .addText((text) =>
        text.setValue(String(s.completedDays)).onChange(async (value) => {
          const days = parseInt(value, 10);
          if (Number.isFinite(days) && days >= 0) {
            s.completedDays = days;
            await this.plugin.saveSettings();
          }
        })
      );
    new Setting(containerEl)
      .setName("Hide empty lists")
      .addToggle((toggle) =>
        toggle.setValue(s.hideEmptyLists).onChange(async (value) => {
          s.hideEmptyLists = value;
          await this.plugin.saveSettings();
          this.plugin.panels().forEach((view) => view.renderLists());
        })
      );

    new Setting(containerEl).setName("Inserting tasks").setHeading();
    this.text(
      "Task template",
      "What a dropped task becomes. Placeholders: {{title}}, {{due}}, {{date}}, {{list}}, {{link}}, {{id}}.",
      "template"
    );
    this.text("Due format", "What {{due}} becomes when the task has a due date. {{date}} is YYYY-MM-DD.", "dueFormat");
    this.text("Link format", "What {{link}} becomes. Placeholders: {{id}}, {{list}}.", "linkFormat");
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

// ── Microsoft sign-in ────────────────────────────────────────────────────────
//
// OAuth 2.0 device code flow against the Microsoft identity platform, with no
// library: the user opens a Microsoft page, enters a short code and signs in,
// while the plugin polls for the tokens. Requests go through Obsidian's
// requestUrl, which is not subject to browser CORS rules.

class NotSignedInError extends Error {}

class MicrosoftAuth {
  constructor(plugin) {
    this.plugin = plugin;
    this.refreshing = null;
    this.session = this.loadSession();
    if (this.session && !this.matchesSettings()) this.signOut();
  }

  get settings() {
    return this.plugin.settings;
  }

  get signedIn() {
    return Boolean(this.session && this.session.refresh_token);
  }

  get account() {
    return this.session ? this.session.account || "" : "";
  }

  // A session belongs to the client ID and tenant it was issued for.
  matchesSettings() {
    return (
      Boolean(this.session) &&
      this.session.client_id === this.settings.clientId.trim() &&
      this.session.tenant === this.tenant()
    );
  }

  tenant() {
    return (this.settings.tenant || "").trim() || "common";
  }

  endpoint(path) {
    return `https://login.microsoftonline.com/${encodeURIComponent(this.tenant())}/oauth2/v2.0/${path}`;
  }

  scope() {
    const asked = (this.settings.scopes || "").split(/[\s,]+/).filter(Boolean);
    return [...new Set([...asked, "offline_access", "openid", "profile"])].join(" ");
  }

  // ── session storage ──

  loadSession() {
    try {
      const store = this.plugin.app.secretStorage;
      const raw = store ? store.getSecret(SESSION_SECRET_ID) : this.plugin.app.loadLocalStorage(SESSION_SECRET_ID);
      return raw ? JSON.parse(raw) : null;
    } catch (_) {
      return null;
    }
  }

  saveSession(session) {
    this.session = session;
    const raw = session ? JSON.stringify(session) : "";
    const store = this.plugin.app.secretStorage;
    // Secret storage has no delete; an empty value reads back as signed out.
    if (store) store.setSecret(SESSION_SECRET_ID, raw);
    else this.plugin.app.saveLocalStorage(SESSION_SECRET_ID, raw || null);
  }

  signOut() {
    this.saveSession(null);
  }

  // ── device code flow ──

  async startDeviceFlow() {
    const res = await postForm(this.endpoint("devicecode"), {
      client_id: this.settings.clientId.trim(),
      scope: this.scope(),
    });
    if (res.status !== 200) throw new Error(describeAuthError(res.json));
    return res.json;
  }

  // Polls until the user finishes signing in. `cancel` is an object whose
  // `cancelled` flag stops the loop, for a closed dialog or an unloaded plugin.
  async completeDeviceFlow(flow, cancel) {
    let interval = Math.max(1, Number(flow.interval) || 5);
    const deadline = Date.now() + (Number(flow.expires_in) || 900) * 1000;
    while (!cancel.cancelled) {
      await sleep(interval * 1000);
      if (cancel.cancelled) break;
      if (Date.now() > deadline) throw new Error("The sign-in code expired. Start again.");
      const res = await postForm(this.endpoint("token"), {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: this.settings.clientId.trim(),
        device_code: flow.device_code,
      });
      if (res.status === 200) {
        this.storeTokens(res.json);
        return;
      }
      const error = res.json && res.json.error;
      if (error === "authorization_pending") continue;
      if (error === "slow_down") {
        interval += 5;
        continue;
      }
      if (error === "authorization_declined") throw new Error("Sign-in was declined.");
      if (error === "expired_token") throw new Error("The sign-in code expired. Start again.");
      throw new Error(describeAuthError(res.json));
    }
    throw new Error("Sign-in cancelled.");
  }

  storeTokens(tokens) {
    const previous = this.session || {};
    const claims = jwtClaims(tokens.id_token);
    this.saveSession({
      client_id: this.settings.clientId.trim(),
      tenant: this.tenant(),
      access_token: tokens.access_token,
      // Microsoft rotates refresh tokens; keep the newest one.
      refresh_token: tokens.refresh_token || previous.refresh_token,
      expires_at: Date.now() + (Number(tokens.expires_in) || 3600) * 1000,
      account: (claims && (claims.preferred_username || claims.email || claims.name)) || previous.account || "",
    });
  }

  // ── access tokens ──

  // A valid access token, refreshed when it is about to expire or when
  // `force` says Graph rejected the current one.
  async accessToken(force = false) {
    if (!this.signedIn) throw new NotSignedInError("Not signed in to Microsoft.");
    const fresh = this.session.access_token && this.session.expires_at - 120000 > Date.now();
    if (fresh && !force) return this.session.access_token;
    // Concurrent callers share one refresh, so a rotated refresh token is
    // never spent twice.
    if (!this.refreshing) {
      this.refreshing = this.refresh().finally(() => {
        this.refreshing = null;
      });
    }
    await this.refreshing;
    return this.session.access_token;
  }

  async refresh() {
    const res = await postForm(this.endpoint("token"), {
      grant_type: "refresh_token",
      client_id: this.settings.clientId.trim(),
      refresh_token: this.session.refresh_token,
      scope: this.scope(),
    });
    if (res.status === 200) {
      this.storeTokens(res.json);
      return;
    }
    const error = res.json && res.json.error;
    if (error === "invalid_grant" || error === "interaction_required" || error === "invalid_client") {
      this.signOut();
      throw new NotSignedInError(`Your Microsoft sign-in has expired. Sign in again. (${describeAuthError(res.json)})`);
    }
    throw new Error(describeAuthError(res.json));
  }
}

class SignInModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this.cancel = { cancelled: false };
  }

  async onOpen() {
    this.setTitle("Sign in to Microsoft");
    if (!this.plugin.settings.clientId.trim()) this.askForApp();
    else this.startSignIn();
  }

  // First sign-in: the app registration's IDs, which the device code needs.
  askForApp() {
    const { contentEl } = this;
    const s = this.plugin.settings;
    contentEl.createDiv({
      cls: "sb-todo-signin-status",
      text: "Enter the IDs of the app you registered in Microsoft Entra. The plugin's README explains how to register one.",
    });
    let clientId = "";
    let tenant = s.tenant || "common";
    new Setting(contentEl)
      .setName("Client ID")
      .setDesc("Application (client) ID")
      .addText((text) => text.setPlaceholder("00000000-0000-0000-0000-000000000000").onChange((v) => (clientId = v.trim())));
    new Setting(contentEl)
      .setName("Tenant")
      .setDesc("Directory (tenant) ID for an app limited to your organisation; otherwise common or consumers.")
      .addText((text) => text.setValue(tenant).onChange((v) => (tenant = v.trim())));
    new Setting(contentEl).addButton((button) =>
      button
        .setButtonText("Continue")
        .setCta()
        .onClick(async () => {
          if (!clientId) {
            new Notice("Enter the client ID.");
            return;
          }
          s.clientId = clientId;
          s.tenant = tenant || "common";
          await this.plugin.saveSettings();
          contentEl.empty();
          this.startSignIn();
        })
    );
  }

  async startSignIn() {
    const { contentEl } = this;
    const status = contentEl.createDiv({ cls: "sb-todo-signin-status", text: "Asking Microsoft for a sign-in code…" });
    let flow;
    try {
      flow = await this.plugin.auth.startDeviceFlow();
    } catch (e) {
      status.setText(`Could not start sign-in: ${e.message}`);
      status.addClass("sb-todo-error");
      return;
    }
    if (this.cancel.cancelled) return;

    status.setText("Open the Microsoft sign-in page, enter this code and sign in. This window closes by itself when you are done.");
    const codeRow = contentEl.createDiv({ cls: "sb-todo-signin-code" });
    codeRow.createEl("code", { text: flow.user_code });
    const actions = contentEl.createDiv({ cls: "sb-todo-signin-actions" });
    const copy = actions.createEl("button", { text: "Copy code" });
    copy.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(flow.user_code);
        copy.setText("Copied");
      } catch (_) {
        new Notice("Could not copy; select the code instead.");
      }
    });
    const open = actions.createEl("button", { cls: "mod-cta", text: "Open sign-in page" });
    open.addEventListener("click", () => window.open(flow.verification_uri, "_blank"));
    contentEl.createDiv({ cls: "sb-todo-signin-hint", text: flow.verification_uri });

    try {
      await this.plugin.auth.completeDeviceFlow(flow, this.cancel);
      if (this.cancel.cancelled) return;
      this.close();
      this.plugin.signedIn();
    } catch (e) {
      if (this.cancel.cancelled) return;
      status.setText(e.message);
      status.addClass("sb-todo-error");
      codeRow.remove();
      actions.remove();
    }
  }

  onClose() {
    this.cancel.cancelled = true;
    this.contentEl.empty();
    if (this.plugin.signInModal === this) this.plugin.signInModal = null;
    // The settings tab, if open, shows the new state next time it renders.
  }
}

// ── Microsoft Graph ──────────────────────────────────────────────────────────

const GRAPH_MAX_CONCURRENT = 3; // To Do throttles hard beyond a few at a time.
const GRAPH_MAX_RETRIES = 5;

async function graphGet(auth, url) {
  let token = await auth.accessToken();
  let retriedAuth = false;
  for (let attempt = 0; ; attempt++) {
    const res = await requestUrl({ url, headers: { Authorization: `Bearer ${token}` }, throw: false });
    if (res.status === 401 && !retriedAuth) {
      retriedAuth = true;
      token = await auth.accessToken(true);
      continue;
    }
    if ([429, 503, 504].includes(res.status) && attempt < GRAPH_MAX_RETRIES) {
      await sleep(retryDelay(res.headers, attempt) * 1000);
      continue;
    }
    if (res.status >= 400) {
      const body = safeJson(res);
      const message = (body && body.error && body.error.message) || `HTTP ${res.status}`;
      if (res.status === 403) throw new Error(`Microsoft refused access to your tasks: ${message}`);
      throw new Error(`Microsoft Graph error: ${message}`);
    }
    return safeJson(res) || {};
  }
}

async function graphGetAll(auth, url) {
  const items = [];
  while (url) {
    const page = await graphGet(auth, url);
    items.push(...(page.value || []));
    url = page["@odata.nextLink"];
  }
  return items;
}

// Every To Do list in To Do's order, each with its open tasks first (by due
// day, undated last, then title) and, when completedDays > 0, the tasks
// completed today or in the days before.
async function fetchTodoLists(auth, completedDays) {
  // One token up front, so parallel requests do not each try to refresh.
  await auth.accessToken();
  const lists = await graphGetAll(auth, `${GRAPH}/me/todo/lists`);
  const filter = taskFilter(completedDays);
  // To Do's tasks endpoint rejects $select, so whole tasks come back.
  const query = `$filter=${encodeURIComponent(filter)}&$top=100`;
  const taskPages = await mapLimit(lists, GRAPH_MAX_CONCURRENT, (list) =>
    graphGetAll(auth, `${GRAPH}/me/todo/lists/${list.id}/tasks?${query}`)
  );
  return lists.map((list, i) => {
    const tasks = taskPages[i]
      .filter((task) => completedDays > 0 || task.status !== "completed")
      .map(taskView);
    tasks.sort(
      (a, b) =>
        a.completed - b.completed ||
        (a.due_date === null) - (b.due_date === null) ||
        (a.due_date || "").localeCompare(b.due_date || "") ||
        a.title.toLowerCase().localeCompare(b.title.toLowerCase())
    );
    return { id: list.id, name: list.displayName || "", wellknown_name: list.wellknownListName || null, tasks };
  });
}

function taskFilter(completedDays) {
  if (completedDays <= 0) return "status ne 'completed'";
  const since = new Date();
  since.setHours(0, 0, 0, 0);
  since.setDate(since.getDate() - completedDays);
  const stamp = since.toISOString().slice(0, 19);
  return `status ne 'completed' or completedDateTime/dateTime ge '${stamp}'`;
}

function taskView(task) {
  const body = task.body || {};
  return {
    id: task.id,
    title: task.title || "",
    importance: task.importance || null,
    // To Do sends a due day as local midnight expressed in UTC.
    due_date: localDayFromUtc(task.dueDateTime),
    completed: task.status === "completed",
    // A completion day comes as a bare date at midnight; take it as is.
    completed_date: task.completedDateTime && task.completedDateTime.dateTime
      ? task.completedDateTime.dateTime.slice(0, 10)
      : null,
    notes: body.contentType === "text" ? (body.content || "").trim() : "",
    categories: task.categories || [],
  };
}

function localDayFromUtc(value) {
  if (!value || !value.dateTime) return null;
  const raw = value.dateTime.replace(/\.\d+$/, "");
  const zone = (value.timeZone || "UTC").toUpperCase();
  const parsed = new Date(zone === "UTC" ? `${raw}Z` : raw);
  if (Number.isNaN(parsed.getTime())) return value.dateTime.slice(0, 10);
  const pad = (n) => String(n).padStart(2, "0");
  return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
}

// ── small helpers ────────────────────────────────────────────────────────────

function postForm(url, fields) {
  return requestUrl({
    url,
    method: "POST",
    contentType: "application/x-www-form-urlencoded",
    body: new URLSearchParams(fields).toString(),
    throw: false,
  }).then((res) => ({ status: res.status, headers: res.headers, json: safeJson(res) }));
}

function safeJson(res) {
  try {
    return res.json;
  } catch (_) {
    return null;
  }
}

// Microsoft's message, trimmed, with a hint for the setup mistakes that
// produce the least obvious errors.
const AUTH_HINTS = {
  AADSTS50059: "Set Tenant to your directory (tenant) ID: this app accepts only accounts in its own organisation.",
  AADSTS700016: "Check the client ID and the tenant: Microsoft cannot find that app there.",
  AADSTS7000218: "In the app registration, turn on Authentication → Allow public client flows.",
  AADSTS65001: "The permission has not been granted yet. An administrator may need to approve the app.",
  AADSTS90002: "Check the tenant setting: Microsoft does not know that directory.",
};

function describeAuthError(body) {
  if (!body) return "No response from Microsoft.";
  const text = String(body.error_description || body.error || "Unknown error.");
  // Microsoft appends trace and correlation IDs; keep the message itself.
  const message = text.split(/\r?\n/)[0].replace(/\s*Trace ID:.*$/, "");
  const code = (message.match(/AADSTS\d+/) || [])[0];
  return code && AUTH_HINTS[code] ? `${message} ${AUTH_HINTS[code]}` : message;
}

function jwtClaims(token) {
  try {
    const part = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = decodeURIComponent(
      atob(part)
        .split("")
        .map((c) => "%" + c.charCodeAt(0).toString(16).padStart(2, "0"))
        .join("")
    );
    return JSON.parse(json);
  } catch (_) {
    return null;
  }
}

function retryDelay(headers, attempt) {
  const header = Object.entries(headers || {}).find(([key]) => key.toLowerCase() === "retry-after");
  const seconds = header ? parseInt(header[1], 10) : NaN;
  return Number.isFinite(seconds) && seconds > 0 ? Math.min(seconds, 60) : Math.min(60, 2 ** (attempt + 1));
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

// Runs fn over items with at most `limit` in flight, keeping results in order.
async function mapLimit(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}
