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

const DEFAULT_SETTINGS = {
  // Run inside WSL through a login shell, so ~/.local/bin is on PATH.
  command: 'wsl.exe -e bash -lc "second-brain todo-lists --json"',
  template: "- [ ] {{title}} {{due}}",
  dueFormat: "📅 {{date}}",
  linkFormat: "https://to-do.office.com/tasks/id/{{id}}/details",
  hideEmptyLists: true,
  timeoutSeconds: 60,
};

module.exports = class TodoPanelPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
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
    return fill(s.template, values)
      .replace(/[ \t]{2,}/g, " ")
      .replace(/[ \t]+$/gm, "");
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
    this.collapsed = new Set();
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
      this.filter = search.value.trim().toLowerCase();
      this.renderLists();
    });
    const refresh = toolbar.createEl("button", { attr: { "aria-label": "Refresh" } });
    setIcon(refresh, "refresh-cw");
    refresh.addEventListener("click", () => this.refresh());

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

  renderLists() {
    this.renderStatus();
    if (!this.listsEl) return;
    this.listsEl.empty();
    const today = localToday();

    for (const list of this.lists) {
      const tasks = list.tasks.filter((task) => this.matches(task, list));
      if (tasks.length === 0 && (this.filter || this.plugin.settings.hideEmptyLists)) continue;

      const details = this.listsEl.createEl("details", { cls: "sb-todo-list" });
      // A filter opens every list that has a match.
      details.open = this.filter ? true : !this.collapsed.has(list.id);
      details.addEventListener("toggle", () => {
        if (this.filter) return;
        if (details.open) this.collapsed.delete(list.id);
        else this.collapsed.add(list.id);
      });
      const summary = details.createEl("summary", { text: list.name });
      summary.createSpan({ cls: "sb-todo-count", text: String(tasks.length) });

      for (const task of tasks) this.renderTask(details, task, list, today);
    }
  }

  renderTask(parent, task, list, today) {
    const row = parent.createDiv({ cls: "sb-todo-task", attr: { draggable: "true" } });
    // A marker per task, with wrapped title lines hanging beside it, so a
    // continuation line never looks like the start of the next task.
    const marker = row.createSpan({ cls: "sb-todo-marker", attr: { "aria-hidden": "true" } });
    if (task.importance === "high") marker.addClass("is-important");

    const body = row.createDiv({ cls: "sb-todo-body" });
    body.createDiv({ cls: "sb-todo-title", text: task.title });

    const meta = [];
    if (task.importance === "high") meta.push(["sb-todo-flag", "Important"]);
    if (task.due_date) {
      const state = task.due_date < today ? " is-overdue" : task.due_date === today ? " is-today" : "";
      meta.push([`sb-todo-due${state}`, shortDate(task.due_date)]);
    }
    if (task.notes) meta.push(["sb-todo-has-notes", "Notes"]);
    if (meta.length) {
      const metaEl = body.createDiv({ cls: "sb-todo-meta" });
      for (const [cls, text] of meta) metaEl.createSpan({ cls, text });
    }
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

function shortDate(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  const sameYear = year === new Date().getFullYear();
  return date.toLocaleDateString([], sameYear ? { day: "numeric", month: "short" } : { day: "numeric", month: "short", year: "numeric" });
}

