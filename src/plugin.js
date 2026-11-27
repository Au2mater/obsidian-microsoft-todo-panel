// Microsoft To Do Panel: a sidebar listing Microsoft To Do tasks. The plugin signs in
// to Microsoft itself (OAuth device code flow, read-only Tasks permission) and
// reads tasks from Microsoft Graph. Dragging a task into a note inserts it as
// a Markdown line, and so does its + button, below the cursor; clicking a
// task the note already holds selects that line.
//
// Plain JavaScript with no dependencies: `npm run build` checks this file and
// writes it unchanged to main.js, the file Obsidian loads (see README.md).

const {
  ItemView,
  MarkdownView,
  Modal,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  moment,
  requestUrl,
  setIcon,
} = require("obsidian");
// Obsidian's own editor modules, for the links beside task lines in notes.
const { Decoration, ViewPlugin, WidgetType } = require("@codemirror/view");
const { RangeSetBuilder, StateEffect } = require("@codemirror/state");

const VIEW_TYPE = "microsoft-todo-panel";
const DRAG_TYPE = "application/x-microsoft-todo-task";
const GRAPH = "https://graph.microsoft.com/v1.0";
// Where the sign-in session lives: Obsidian's secret storage, kept on this
// device and outside the vault, so it is never synced or committed. The ID
// carries a hash of the vault name, so vaults signed in to different apps do
// not overwrite each other. Versions up to 1.0.3 used the bare legacy ID.
const SESSION_SECRET_PREFIX = "todo-panel-microsoft-session";
// The last fetched lists, in this vault's local storage on this device, so the
// panel can show them at once on startup while it refreshes.
const LIST_CACHE_KEY = "microsoft-todo-panel-lists";
// Recent searches, newest first, on this device like Obsidian's own.
const SEARCH_HISTORY_KEY = "microsoft-todo-panel-search-history";
const SEARCH_HISTORY_MAX = 10;
// Stands for the active note's name in a search, as {{title}} does in
// Obsidian's templates: "list:{{title}}" finds the list named after the note.
const TITLE_PLACEHOLDER = "{{title}}";
// The search left in the box, so it is still there after a restart.
const SEARCH_KEY = "microsoft-todo-panel-search";
// The search box's operators, offered in its menu as Obsidian's search offers
// path:, file: and tag:. Those with values offer them once typed.
// `takesTitle` marks those that also offer TITLE_PLACEHOLDER as a value.
const SEARCH_OPERATORS = [
  { op: "title:", info: "match the task's title", takesTitle: true },
  { op: "note:", info: "match the task's notes", takesTitle: true },
  { op: "list:", info: "match the list's name", takesTitle: true },
  {
    op: "due-date:",
    info: "today, overdue, none or a date",
    values: [
      ["today", "due today"],
      ["tomorrow", "due tomorrow"],
      ["overdue", "due before today"],
      ["upcoming", "due after today"],
      ["none", "no due date"],
    ],
  },
  {
    op: "is:",
    info: "open, completed or important",
    values: [
      ["open", "not completed"],
      ["completed", "completed tasks, when shown"],
      ["important", "flagged as important"],
    ],
  },
];

const DEFAULT_SETTINGS = {
  // Microsoft sign-in. There is no shared app registration: each user
  // registers their own (see README.md) and pastes its IDs here.
  clientId: "",
  tenant: "common",
  scopes: "Tasks.Read",
  completedDays: 7,
  template: "- [ ] {{title}} {{due}}",
  dueFormat: "📅 {{date}}",
  // How daily note file names spell their date (Moment.js format).
  dailyNoteFormat: "YYYY-MM-DD",
  // Panel state, kept between sessions.
  tab: "lists",
  showCompleted: false,
  onlyInNote: false,
  collapsed: { lists: [], dates: ["date:none"] },
};

// Settings that earlier versions had: an external command (before 0.2), and
// options folded into fixed behaviour in 1.2.
const RETIRED_SETTINGS = ["command", "timeoutSeconds", "linkFormat", "hideEmptyLists"];
// What {{link}} in the task template becomes.
// How long a task shown from a note flashes; matches styles.css.
const REVEAL_FLASH_MS = 1600;
const TASK_LINK_FORMAT = "https://to-do.office.com/tasks/id/{{id}}/details";

module.exports = class TodoPanelPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    for (const key of RETIRED_SETTINGS) delete this.settings[key];
    this.auth = new MicrosoftAuth(this);
    this.registerView(VIEW_TYPE, (leaf) => new TodoPanelView(leaf, this));
    this.addRibbonIcon("list-checks", "Open Microsoft To Do panel", () => this.openPanel());
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
    this.settingTab = new TodoPanelSettingTab(this.app, this);
    this.addSettingTab(this.settingTab);

    // Keep the "in note" markers in step with the note being worked on.
    // Opening a note fires several events at once, and typing fires one per
    // keystroke, so all of them share one short timer and the note is checked
    // once per burst.
    this.registerEvent(this.app.workspace.on("active-leaf-change", () => this.scheduleNoteCheck(50)));
    this.registerEvent(this.app.workspace.on("file-open", () => this.scheduleNoteCheck(50)));
    this.registerEvent(this.app.workspace.on("editor-change", () => this.scheduleNoteCheck(300)));

    // A task dropped into a note lands on a line of its own. The editor
    // inserts the dragged text at the drop point itself; right after, the
    // task is moved off any line it was dropped into the middle of.
    this.registerEvent(
      this.app.workspace.on("editor-drop", (evt, editor) => {
        const dt = evt.dataTransfer;
        if (!dt || !Array.from(dt.types || []).includes(DRAG_TYPE)) return;
        const text = dt.getData("text/plain");
        window.setTimeout(() => placeOnOwnLine(editor, text), 0);
      })
    );
    this.registerEvent(
      this.app.vault.on("modify", (file) => {
        if (file === this.app.workspace.getActiveFile()) this.scheduleNoteCheck(300);
      })
    );

    // A link beside each note line that holds a panel task, which shows the
    // task in the panel. A visual layer only: the note's text is untouched.
    this.taskIndex = null;
    this.taskIndexDirty = true;
    this.taskShape = this.taskLineShape();
    this.taskLinkViews = new Set();
    this.registerEditorExtension(taskLinkExtension(this));

    // "Today" and overdue colours move on at midnight even if nothing else
    // happens in the panel.
    this.today = localToday();
    this.registerInterval(
      window.setInterval(() => {
        if (localToday() === this.today) return;
        this.today = localToday();
        this.panels().forEach((view) => view.renderLists());
      }, 60 * 1000)
    );
  }

  onunload() {
    window.clearTimeout(this.noteCheckTimer);
    window.clearTimeout(this.taskLinkTimer);
    this.taskLinkViews.clear();
    // Stops a sign-in that is still waiting for the user.
    if (this.signInModal) this.signInModal.close();
  }

  scheduleNoteCheck(delay) {
    window.clearTimeout(this.noteCheckTimer);
    this.noteCheckTimer = window.setTimeout(() => {
      this.panels().forEach((view) => view.updateNoteMatches());
    }, delay);
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
      if (!leaf) return;
      await leaf.setViewState({ type: VIEW_TYPE, active: true });
    }
    await this.app.workspace.revealLeaf(leaf);
  }

  async saveSettings() {
    const shape = this.taskLineShape();
    if (shape !== this.taskShape) this.tasksChanged();
    this.taskShape = shape;
    await this.saveData(this.settings);
  }

  // The settings that decide what a task line looks like: the template and
  // the due date format.
  taskLineShape() {
    return `${this.settings.template}\u0000${this.settings.dueFormat}`;
  }

  // ── links from note lines to the panel ──

  // The panel's tasks changed: the links in open notes are redrawn, once,
  // after whatever changed them has finished.
  tasksChanged() {
    this.taskIndexDirty = true;
    window.clearTimeout(this.taskLinkTimer);
    this.taskLinkTimer = window.setTimeout(() => {
      for (const links of this.taskLinkViews) links.redraw();
    }, 0);
  }

  // What note lines are matched against, built on first use after a change
  // (see buildTaskIndex): the panel's tasks, or, while no open panel holds
  // any (it is closed, or not loaded yet in a hidden sidebar tab), the last
  // fetched ones. Null when there are none, as when signed out.
  taskIndexNow() {
    if (this.taskIndexDirty) {
      this.taskIndexDirty = false;
      const view = this.panels().find((panel) => panel.lists.length);
      const lists = view ? view.lists : (this.loadListCache() || { lists: [] }).lists;
      this.taskIndex = buildTaskIndex(lists, (task, list) => this.renderTask(task, list).trim());
    }
    return this.taskIndex;
  }

  // Clicking a link in a note: shows that task in the panel.
  async revealTask(id) {
    await this.openPanel();
    const view = this.panels()[0];
    if (view) view.reveal(id);
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
    this.saveListCache(null);
    this.tasksChanged();
    this.panels().forEach((view) => view.clear());
    this.refreshSettingTab();
    new Notice("Signed out of Microsoft.");
  }

  // Called by the sign-in dialog once Microsoft has issued tokens.
  signedIn() {
    const who = this.auth.account ? ` as ${this.auth.account}` : "";
    new Notice(`Signed in to Microsoft${who}.`);
    this.panels().forEach((view) => view.refresh());
    this.refreshSettingTab();
  }

  // Redraws the settings tab if it is open, so it shows the new sign-in state.
  refreshSettingTab() {
    if (this.settingTab && this.settingTab.containerEl.isConnected) this.settingTab.display();
  }

  // ── list cache ──

  // Recent searches, newest first, shared by every panel in this vault.
  searchHistory() {
    try {
      const saved = this.app.loadLocalStorage(SEARCH_HISTORY_KEY);
      return Array.isArray(saved) ? saved.filter((q) => typeof q === "string") : [];
    } catch (e) {
      return [];
    }
  }

  rememberSearch(query) {
    query = String(query || "").trim();
    if (!query) return;
    const history = [query, ...this.searchHistory().filter((q) => q !== query)].slice(0, SEARCH_HISTORY_MAX);
    this.app.saveLocalStorage(SEARCH_HISTORY_KEY, history);
  }

  savedSearch() {
    try {
      const saved = this.app.loadLocalStorage(SEARCH_KEY);
      return typeof saved === "string" ? saved : "";
    } catch (e) {
      return "";
    }
  }

  saveSearch(query) {
    this.app.saveLocalStorage(SEARCH_KEY, query || null);
  }

  clearSearchHistory() {
    this.app.saveLocalStorage(SEARCH_HISTORY_KEY, null);
  }

  loadListCache() {
    try {
      const cached = this.app.loadLocalStorage(LIST_CACHE_KEY);
      if (!cached || !Array.isArray(cached.lists)) return null;
      // Only the account that fetched them may see them.
      if (!this.auth.signedIn || cached.account !== this.auth.account) return null;
      return cached;
    } catch (_) {
      return null;
    }
  }

  saveListCache(lists) {
    try {
      this.app.saveLocalStorage(
        LIST_CACHE_KEY,
        lists ? { account: this.auth.account, fetchedAt: Date.now(), lists } : null
      );
    } catch (_) {
      // A full or unavailable local storage only costs the instant start.
    }
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
      link: fill(TASK_LINK_FORMAT, { id: encodeURIComponent(task.id), list: encodeURIComponent(list.id) }),
    };
    let line = fill(s.template, values);
    if (task.completed) line = line.replace("[ ]", "[x]");
    return line
      .replace(/[ \t]{2,}/g, " ")
      .replace(/[ \t]+$/gm, "");
  }

  // The date a daily note stands for, as YYYY-MM-DD, or null for any other
  // note, read from the file name with the "Daily note date format" setting.
  dailyNoteDate(file) {
    if (!file || !file.basename) return null;
    const format = (this.settings.dailyNoteFormat || "").trim() || "YYYY-MM-DD";
    // A format with folders ("YYYY/MM/YYYY-MM-DD") names the path, so match
    // its last part against the file name.
    const nameFormat = format.split("/").pop();
    const parsed = moment(file.basename, nameFormat, true);
    return parsed.isValid() ? parsed.format("YYYY-MM-DD") : null;
  }

  // The note being worked in: the active Markdown view, or, while focus is
  // in the sidebar (clicking the panel moves it there), the Markdown view
  // that was active last in the main area.
  noteView() {
    const active = this.app.workspace.getActiveViewOfType(MarkdownView);
    if (active) return active;
    const recent = this.app.workspace.getMostRecentLeaf();
    return recent && recent.view instanceof MarkdownView ? recent.view : null;
  }

  // The note being worked in, as { file, text }, or null.
  activeNote() {
    const view = this.noteView();
    return view && view.file ? { file: view.file, text: view.getViewData() } : null;
  }

  // The note to act on, switched out of reading view and made active, so its
  // editor can select or insert. Null, with a notice, when no note is open.
  async noteEditor() {
    const view = this.noteView();
    if (!view) {
      new Notice("Open a note first.");
      return null;
    }
    if (view.getMode() === "preview") await view.setState({ ...view.getState(), mode: "source" }, { history: false });
    this.app.workspace.setActiveLeaf(view.leaf, { focus: true });
    return view.editor;
  }

  // Clicking a task the note already holds selects that line and scrolls to
  // it. Clicking any other task does nothing; its + button adds it.
  async showInNote(task, list) {
    const view = this.noteView();
    if (!view) return;
    // Read the note as it is now, not as the markers last saw it.
    const line = findTaskLine(noteTaskLines(view.getViewData()), normalizedTitle(task), this.renderTask(task, list).trim());
    if (line < 0) return;
    const editor = await this.noteEditor();
    if (!editor) return;
    const from = { line, ch: 0 };
    const to = { line, ch: editor.getLine(line).length };
    editor.setSelection(from, to);
    editor.scrollIntoView({ from, to }, true);
    editor.focus();
  }

  // The + button: adds the task on a new line below the cursor's line,
  // indented like it.
  async addToNote(task, list) {
    const view = this.noteView();
    const editor = await this.noteEditor();
    if (!editor) return;
    // Never inside the note's properties: below them when the cursor is in
    // them (as it is in a note opened but not yet clicked into).
    const cache = view && view.file ? this.app.metadataCache.getFileCache(view.file) : null;
    const properties = cache && cache.frontmatterPosition;
    let line = editor.getCursor().line;
    if (properties && line <= properties.end.line) line = properties.end.line;
    insertOnOwnLine(editor, this.renderTask(task, list).trim(), line);
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
    this.refreshQueued = false;
    this.error = null;
    this.fetchedAt = null;
    // The note text the "in note" markers were last computed from.
    this.checkedNoteKey = null;
  }

  // The tasks shown. Every change redraws the links in notes, which match
  // note lines against these tasks.
  get lists() {
    return this.shownLists;
  }

  set lists(lists) {
    this.shownLists = lists;
    this.plugin.tasksChanged();
  }

  getViewType() {
    return VIEW_TYPE;
  }

  getDisplayText() {
    return "Microsoft To Do";
  }

  getIcon() {
    return "list-checks";
  }

  async onOpen() {
    const root = this.contentEl;
    root.empty();
    root.addClass("sb-todo-panel");

    const toolbar = root.createDiv({ cls: "sb-todo-toolbar" });
    const search = (this.searchEl = toolbar.createEl("input", {
      type: "search",
      attr: { placeholder: "Search..." },
    }));
    // A search opens every group with a match, using its own collapsed
    // state, so clearing it brings the groups back as they were.
    const apply = (remember = false) => {
      // Case is kept for OR; the terms themselves ignore it.
      const filter = search.value.trim();
      if (filter !== this.filter) {
        this.filter = filter;
        this.filterCollapsed = new Set();
        this.plugin.saveSearch(filter);
        this.renderLists();
      }
      if (remember) this.plugin.rememberSearch(search.value);
    };
    this.searchMenu = new SearchMenu(search, {
      history: () => this.plugin.searchHistory(),
      clearHistory: () => this.plugin.clearSearchHistory(),
      noteName: () => this.noteName,
      apply,
    });
    // The search left in the box last time, until it is cleared.
    search.value = this.filter = this.plugin.savedSearch();
    search.addEventListener("input", () => {
      apply();
      this.searchMenu.update();
    });
    search.addEventListener("focus", () => this.searchMenu.update());
    // A click in a box already focused, say after Escape, opens it again.
    // Moving the cursor also changes the word the menu would replace.
    search.addEventListener("click", () => this.searchMenu.update());
    search.addEventListener("keyup", (evt) => {
      if (["ArrowLeft", "ArrowRight", "Home", "End"].includes(evt.key)) this.searchMenu.update();
    });
    search.addEventListener("keydown", (evt) => {
      if (this.searchMenu.onKeyDown(evt)) return;
      if (evt.key === "Enter") {
        apply(true);
        this.searchMenu.close();
      }
    });
    // Leaving the box keeps what was searched for, as Obsidian does.
    search.addEventListener("blur", () => {
      this.searchMenu.close();
      apply(true);
    });
    this.completedEl = toolbar.createEl("button", { cls: "sb-todo-toggle" });
    this.completedEl.addEventListener("click", () => {
      this.plugin.settings.showCompleted = !this.plugin.settings.showCompleted;
      this.plugin.saveSettings();
      this.plugin.panels().forEach((view) => view.renderLists());
    });
    this.inNoteEl = toolbar.createEl("button", { cls: "sb-todo-toggle" });
    this.inNoteEl.addEventListener("click", () => {
      this.plugin.settings.onlyInNote = !this.plugin.settings.onlyInNote;
      this.plugin.saveSettings();
      this.plugin.panels().forEach((view) => view.renderLists());
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
    // Only the lists scroll; the controls above them stay in place.
    this.listsEl = root.createDiv({ cls: "sb-todo-lists" });

    // Show the lists from last time straight away, then refresh them.
    const cached = this.plugin.loadListCache();
    if (cached) {
      this.lists = cached.lists;
      this.fetchedAt = new Date(cached.fetchedAt);
    }
    this.renderLists();

    // Fetching takes a few seconds. Obsidian waits for onOpen while it
    // restores the workspace, so the fetch starts only once the layout is
    // ready and the app is idle, and is never awaited here: startup and
    // editing stay unblocked.
    this.app.workspace.onLayoutReady(() => {
      const start = () => this.refresh();
      if (window.requestIdleCallback) window.requestIdleCallback(start, { timeout: 3000 });
      else window.setTimeout(start, 500);
    });
  }

  async refresh() {
    // A refresh asked for while one runs (say, right after signing in again)
    // runs once the current one ends, rather than being dropped.
    if (this.loading) {
      this.refreshQueued = true;
      return;
    }
    if (!this.plugin.auth.signedIn) {
      this.clear();
      return;
    }
    const auth = this.plugin.auth;
    const epoch = auth.epoch;
    this.loading = true;
    this.error = null;
    this.renderStatus();
    try {
      const lists = await this.plugin.fetchLists();
      // Signed out, or into another account, while this was loading: these
      // lists belong to a session that no longer exists.
      if (auth.epoch === epoch) {
        this.lists = lists;
        this.fetchedAt = new Date();
        this.plugin.saveListCache(lists);
      }
    } catch (e) {
      // An expired sign-in ends the session itself, which also changes the
      // epoch; its message still needs showing, and its cached lists go.
      if (e.expired) this.plugin.saveListCache(null);
      if (auth.epoch === epoch || e.expired) this.error = e.message;
    } finally {
      this.loading = false;
      if (!auth.signedIn) this.lists = [];
      this.renderLists();
    }
    if (this.refreshQueued) {
      this.refreshQueued = false;
      this.refresh();
    }
  }

  // Back to the signed-out state: no tasks, and a way to sign in.
  clear() {
    this.lists = [];
    this.fetchedAt = null;
    this.error = null;
    this.checkedNoteKey = null;
    this.renderLists();
  }

  // The links in notes fall back to the last fetched tasks.
  async onClose() {
    if (this.searchMenu) this.searchMenu.close();
    this.lists = [];
  }

  // Shows one task: brings back a row that the filter or a toggle hides,
  // opens its group, scrolls to it and flashes it.
  reveal(id) {
    // A flash still running for another task stops; redraws below skip it.
    this.revealed = null;
    let row = this.rowsById && this.rowsById.get(id);
    if (!row) {
      const list = this.lists.find((l) => l.tasks.some((t) => t.id === id));
      if (!list) {
        new Notice("That task is no longer in Microsoft To Do.");
        return;
      }
      const task = list.tasks.find((t) => t.id === id);
      // Undo whatever hides it, one thing at a time, least disruptive first;
      // only what hides this task changes.
      const s = this.plugin.settings;
      let settingsChanged = false;
      const undo = [
        () => this.filter && !this.matches(task, list) && ((this.filter = ""), this.searchEl && (this.searchEl.value = ""), this.plugin.saveSearch(""), true),
        () => task.completed && !s.showCompleted && ((s.showCompleted = true), (settingsChanged = true)),
        () => s.onlyInNote && ((s.onlyInNote = false), (settingsChanged = true)),
      ];
      for (const step of undo) {
        if (!step()) continue;
        this.renderLists();
        row = this.rowsById.get(id);
        if (row) break;
      }
      if (settingsChanged) {
        this.plugin.saveSettings();
        // The toggles are shared: other panels show the change too.
        for (const other of this.plugin.panels()) if (other !== this) other.renderLists();
      }
      if (!row) return;
    }
    // Opened as if by hand, so it stays open.
    const group = row.closest("details");
    if (group && !group.open) group.open = true;
    this.revealed = { id, at: Date.now() };
    this.showRevealed();
  }

  // Scrolls to the task shown last and flashes it. A redraw during the flash
  // (a panel that just opened refreshes straight away) rebuilds the rows, so
  // renderLists calls this again and the flash carries on where it was.
  showRevealed() {
    const revealed = this.revealed;
    if (!revealed) return;
    const elapsed = Date.now() - revealed.at;
    if (elapsed >= REVEAL_FLASH_MS) {
      this.revealed = null;
      return;
    }
    const row = this.rowsById && this.rowsById.get(revealed.id);
    if (!row) return;
    row.scrollIntoView({ block: "center" });
    row.style.animationDelay = `-${elapsed}ms`;
    row.removeClass("is-revealed");
    // Restarts the flash when the same task is shown twice in a row.
    void row.offsetWidth;
    row.addClass("is-revealed");
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
      this.statusEl.setText(this.fetchedAt ? "Refreshing tasks…" : "Loading tasks…");
    } else if (this.error) {
      this.statusEl.createDiv({ cls: "sb-todo-error", text: this.error });
    } else if (this.fetchedAt) {
      const time = this.fetchedAt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
      this.statusEl.setText(`Updated ${time} · drag a task into a note, or use +`);
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
    // Only the groups on screen change; hidden ones (filtered out, or lists
    // with nothing to show) keep their state.
    for (const g of groups) {
      if (anyOpen) collapsed.add(g.key);
      else collapsed.delete(g.key);
    }
    this.setCollapsed(collapsed);
    this.renderLists();
  }

  // The groups the current tab shows, each with its matching tasks in order.
  groups() {
    const showCompleted = this.plugin.settings.showCompleted;
    const onlyInNote = this.plugin.settings.onlyInNote;
    // Open tasks first in every group; completed ones, when shown, follow.
    const openFirst = (items) => [
      ...items.filter((e) => !e.task.completed),
      ...items.filter((e) => e.task.completed),
    ];
    const entries = (list) =>
      list.tasks
        .filter(
          (task) =>
            (showCompleted || !task.completed) &&
            (!onlyInNote || this.inNote(task, list)) &&
            this.matches(task, list)
        )
        .map((task) => ({ task, list }));

    if (this.tab === "lists") {
      return this.lists
        .map((list) => ({ key: list.id, name: list.name, items: openFirst(entries(list)) }))
        // Lists with nothing to show only take up room.
        .filter((g) => g.items.length > 0);
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
    return groups.filter((g) => g.items.length > 0 || (!this.filter && !onlyInNote && g.key !== "date:none"));
  }

  renderLists() {
    this.renderStatus();
    if (!this.listsEl) return;
    this.listsEl.empty();
    this.rowsById = new Map();
    if (this.completedEl) {
      const on = this.plugin.settings.showCompleted;
      const days = Number(this.plugin.settings.completedDays) || 0;
      // With no completed tasks loaded the toggle would do nothing.
      this.completedEl.toggleClass("is-hidden", days === 0);
      this.completedEl.toggleClass("is-active", on);
      this.completedEl.setAttr(
        "aria-label",
        on ? "Hide completed tasks" : `Show completed tasks (last ${days} day${days === 1 ? "" : "s"})`
      );
      this.completedEl.setAttr("aria-pressed", String(on));
      setIcon(this.completedEl, "circle-check");
    }
    if (this.inNoteEl) {
      const on = this.plugin.settings.onlyInNote;
      this.inNoteEl.toggleClass("is-active", on);
      this.inNoteEl.setAttr("aria-label", on ? "Show all tasks" : "Show only tasks in the active note");
      this.inNoteEl.setAttr("aria-pressed", String(on));
      setIcon(this.inNoteEl, "file-check");
    }
    for (const [key, el] of Object.entries(this.tabEls || {})) {
      el.toggleClass("is-active", key === this.tab);
      el.setAttr("aria-selected", String(key === this.tab));
    }

    const today = localToday();
    const groups = this.groups();
    const collapsed = this.collapsedSet();
    this.inNoteSignature = this.inNoteIds();
    if (this.plugin.settings.onlyInNote && groups.length === 0 && this.lists.length) {
      const text = !this.noteLines
        ? "Open a note to see the tasks it holds."
        : this.filter
          ? "No task in this note matches the search."
          : this.inNoteIds()
            ? "This note's tasks are all completed; show completed tasks to see them."
            : "None of your tasks are in this note yet.";
      this.listsEl.createDiv({ cls: "sb-todo-empty", text });
    }

    for (const group of groups) {
      const details = this.listsEl.createEl("details", { cls: "sb-todo-list" });
      details.open = !collapsed.has(group.key);
      details.addEventListener("toggle", () => {
        const set = this.collapsedSet();
        // Setting `open` from code also fires "toggle", one per group on every
        // render; only a real change by the user is saved.
        if (details.open === !set.has(group.key)) return;
        if (details.open) set.delete(group.key);
        else set.add(group.key);
        this.setCollapsed(set);
        this.updateToggleAll();
      });
      const summary = details.createEl("summary", { text: group.name });
      const done = group.items.filter((e) => e.task.completed).length;
      summary.createSpan({ cls: "sb-todo-count", text: String(group.items.length - done) });
      if (done) summary.createSpan({ cls: "sb-todo-done-count", text: `+${done} done` });
      // Filled in by updateNoteMatches: how many of the group's tasks the
      // active note holds, visible whether the group is open or collapsed.
      summary.createSpan({ cls: "sb-todo-group-in-note" });
      if (group.items.length === 0) {
        details.createDiv({ cls: "sb-todo-empty", text: "Nothing here" });
      }
      for (const { task, list } of group.items) {
        this.renderTask(details, task, list, today, this.tab === "dates");
      }
    }
    for (const el of this.listsEl.querySelectorAll(".sb-todo-list")) this.markGroupNoteMatches(el);
    this.updateToggleAll();
    this.showRevealed();
    // Rows were marked against the note as last read; this re-reads it only
    // if the active note changed since.
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
    appendLinkedText(body.createDiv({ cls: "sb-todo-title" }), task.title, this.plugin.app);

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
    const metaEl = body.createDiv({ cls: "sb-todo-meta" });
    // Shown by CSS only while the active note already has this task.
    metaEl.createSpan({ cls: "sb-todo-in-note", text: "In note" });
    for (const [cls, text] of meta) metaEl.createSpan({ cls, text });
    if (task.notes) this.renderNotes(body, metaEl, task);
    // A deliberate button to add the task, so a stray click never edits a
    // note. Hidden while the note already holds the task.
    const add = row.createEl("button", { cls: "sb-todo-add", attr: { "aria-label": "Add to note" } });
    setIcon(add, "plus");
    add.addEventListener("click", (evt) => {
      evt.stopPropagation();
      this.plugin.addToNote(task, list).catch((e) => new Notice(`Could not add the task: ${e.message}`));
    });
    row.sbEntry = { task, list };
    this.rowsById.set(task.id, row);
    this.markNoteMatch(row);

    row.addEventListener("dragstart", (evt) => {
      const markdown = this.plugin.renderTask(task, list);
      // The editor inserts text/plain at the drop point on its own.
      evt.dataTransfer.setData("text/plain", markdown);
      evt.dataTransfer.setData(DRAG_TYPE, JSON.stringify({ id: task.id, list_id: list.id }));
      row.addClass("is-dragging");
    });
    row.addEventListener("dragend", () => row.removeClass("is-dragging"));
    // A task the note already holds is shown there.
    row.addEventListener("click", () => {
      this.plugin.showInNote(task, list).catch((e) => new Notice(`Could not show the task: ${e.message}`));
    });
  }

  // A "Notes" toggle in the meta line that shows the task's notes under it.
  // Which tasks have their notes open is kept while the panel is open.
  renderNotes(body, metaEl, task) {
    this.openNotes = this.openNotes || new Set();
    const toggle = metaEl.createEl("button", { cls: "sb-todo-notes-toggle", text: "Notes" });
    const notes = body.createDiv({ cls: "sb-todo-notes" });
    appendLinkedText(notes, task.notes, this.plugin.app);
    const show = (open) => {
      notes.toggleClass("is-hidden", !open);
      toggle.toggleClass("is-active", open);
      toggle.setAttr("aria-expanded", String(open));
    };
    show(this.openNotes.has(task.id));
    toggle.addEventListener("click", (evt) => {
      // Not a click on the task itself, which would insert or show it.
      evt.stopPropagation();
      const open = !this.openNotes.has(task.id);
      if (open) this.openNotes.add(task.id);
      else this.openNotes.delete(task.id);
      show(open);
    });
    // Selecting text in the notes should not insert the task either.
    notes.addEventListener("click", (evt) => evt.stopPropagation());
  }

  // Re-reads the active note and marks the tasks it already contains,
  // without rebuilding the panel, so scrolling and open groups stay put.
  async updateNoteMatches() {
    const note = await this.plugin.activeNote();
    // Typing, saving and switching tabs fire often; skip the work when the
    // note and the settings that shape matching are what they were.
    const s = this.plugin.settings;
    const key = note
      ? [note.file ? note.file.path : "", s.template, s.dueFormat, s.dailyNoteFormat, note.text].join("\u0000")
      : "";
    if (key === this.checkedNoteKey) return;
    this.checkedNoteKey = key;
    this.noteLines = note ? noteTaskLines(note.text) : null;
    const noteName = note && note.file ? note.file.basename : "";
    const renamed = noteName !== this.noteName;
    this.noteName = noteName;
    const refDate = (note && this.plugin.dailyNoteDate(note.file)) || null;
    if (refDate !== this.refDate) {
      this.refDate = refDate;
      // Sections depend on the day; rebuilding calls back here once more,
      // when the day no longer differs.
      if (this.tab === "dates") return this.renderLists();
    }
    if (!this.listsEl) return;
    // A search for {{title}} now means another note.
    if (renamed && this.filter.toLowerCase().includes(TITLE_PLACEHOLDER)) return this.renderLists();
    // Showing only the note's tasks: rebuild when that set of tasks changed.
    if (this.plugin.settings.onlyInNote && this.inNoteIds() !== this.inNoteSignature) return this.renderLists();
    for (const row of this.listsEl.querySelectorAll(".sb-todo-task")) this.markNoteMatch(row);
    for (const group of this.listsEl.querySelectorAll(".sb-todo-list")) this.markGroupNoteMatches(group);
  }

  // The IDs of the tasks the active note holds, as one string to compare.
  inNoteIds() {
    if (!this.noteLines) return "";
    const ids = [];
    for (const list of this.lists) for (const task of list.tasks) if (this.inNote(task, list)) ids.push(task.id);
    return ids.join(",");
  }

  // "○2 ✓1" beside a group's title: its tasks in the active note, open and
  // ticked off. Empty, and so hidden, when the note holds none of them.
  markGroupNoteMatches(group) {
    const label = group.querySelector(".sb-todo-group-in-note");
    if (!label) return;
    const checked = group.querySelectorAll(".sb-todo-task.is-checked-in-note").length;
    const open = group.querySelectorAll(".sb-todo-task.is-in-note").length - checked;
    label.empty();
    if (open) label.createSpan({ cls: "is-open", text: `○${open}` });
    if (checked) label.createSpan({ cls: "is-checked", text: `✓${checked}` });
    const where = this.noteName || "the active note";
    const parts = [open && `${open} open`, checked && `${checked} ticked off`].filter(Boolean);
    label.setAttr("title", parts.length ? `In ${where}: ${parts.join(", ")}` : "");
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
    const title = normalizedTitle(task);
    if (title && lines.titles.has(title)) return lines.titles.get(title).checked ? "checked" : "open";
    const rendered = this.plugin.renderTask(task, list).trim();
    return lines.raw.has(rendered) ? "open" : null;
  }

  matches(task, list) {
    if (!this.filter) return true;
    if (this.parsedFilter !== this.filter) {
      this.parsedFilter = this.filter;
      this.searchTerms = parseSearch(this.filter);
    }
    return taskMatches(this.searchTerms, task, list, localToday(), this.noteName);
  }
}

// ── Search ───────────────────────────────────────────────────────────────────

// Splits a search into alternatives at each OR (in capitals, as in Obsidian),
// each a list of terms that must all match: plain words, "quoted phrases",
// operator:value pairs (value quoted or not), each negated by a leading "-".
// An unknown operator is just a word, so "http://x" still finds itself.
function parseSearch(text) {
  const groups = [[]];
  const tokens = String(text).match(/(?:[^\s"]+|"[^"]*"?)+/g) || [];
  for (const token of tokens) {
    if (token === "OR") {
      groups.push([]);
      continue;
    }
    const terms = groups[groups.length - 1];
    let rest = token;
    // A lone "-" is a term not yet typed, so it matches nothing yet.
    const negated = rest.startsWith("-");
    if (negated) rest = rest.slice(1);
    const m = /^([a-z-]+:)(.*)$/i.exec(rest);
    const known = m && SEARCH_OPERATORS.some((o) => o.op === m[1].toLowerCase());
    const op = known ? m[1].toLowerCase() : null;
    const value = (known ? m[2] : rest).replace(/"/g, "").toLowerCase();
    if (value) terms.push({ op, value, negated });
  }
  // A dangling OR ("milk OR") adds nothing rather than matching everything.
  const kept = groups.filter((g) => g.length);
  return kept.length ? kept : [[]];
}

// Whether a task matches any of the search's alternatives. `noteName` is
// what {{title}} stands for; with no note open, a term using it matches
// nothing (and so, negated, everything).
function taskMatches(groups, task, list, today, noteName = "") {
  const due = task.due_date || "";
  const has = (text, value) => String(text || "").toLowerCase().includes(value);
  const test = ({ op, value }) => {
    if (value.includes(TITLE_PLACEHOLDER)) {
      if (!noteName) return false;
      value = value.split(TITLE_PLACEHOLDER).join(noteName.toLowerCase());
    }
    switch (op) {
      case "title:":
        return has(task.title, value);
      case "note:":
        return has(task.notes, value);
      case "list:":
        return has(list.name, value);
      case "due-date:":
        if (value === "today") return due === today;
        if (value === "tomorrow") return due === addDays(today, 1);
        if (value === "overdue") return !!due && due < today && !task.completed;
        if (value === "upcoming") return due > today;
        if (value === "none") return !due;
        // A date, or the start of one: 2026-11 is any day in November.
        return !!due && due.startsWith(value);
      case "is:":
        if (value === "open") return !task.completed;
        if (value === "completed" || value === "done") return !!task.completed;
        if (value === "important") return task.importance === "high";
        return false;
      default:
        return has(`${task.title} ${task.notes || ""} ${list.name}`, value);
    }
  };
  return groups.some((terms) => terms.every((term) => test(term) !== term.negated));
}

function addDays(isoDate, days) {
  const [year, month, day] = isoDate.split("-").map(Number);
  return isoDay(new Date(year, month - 1, day + days));
}

// What the search menu offers for the text before the cursor: operators
// while a word is starting, an operator's values once it is typed, and the
// history while the box is empty. `token` is the word being replaced.
// `noteName` is shown beside {{title}}, as what it stands for now.
function searchSuggestions(before, history, noteName = "") {
  const token = /\S*$/.exec(before)[0];
  const sign = token.startsWith("-") ? "-" : "";
  const word = token.slice(sign.length).toLowerCase();
  const options = [];
  const withValue = SEARCH_OPERATORS.find((o) => (o.values || o.takesTitle) && word.startsWith(o.op));
  const titleInfo = noteName ? `active note: ${noteName}` : "the active note's name";
  if (withValue) {
    const typed = word.slice(withValue.op.length);
    const values = withValue.takesTitle ? [[TITLE_PLACEHOLDER, titleInfo]] : withValue.values;
    for (const [value, info] of values) {
      if (value.startsWith(typed) && value !== typed) options.push({ text: `${sign}${withValue.op}${value} `, label: value, info });
    }
  } else if (word.startsWith("{")) {
    // A bare {{title}}: the note's name anywhere in the task.
    if (TITLE_PLACEHOLDER.startsWith(word) && word !== TITLE_PLACEHOLDER) {
      options.push({ text: `${sign}${TITLE_PLACEHOLDER} `, label: TITLE_PLACEHOLDER, info: titleInfo });
    }
  } else if (!word.includes(":")) {
    for (const o of SEARCH_OPERATORS) {
      if (o.op.startsWith(word)) options.push({ text: sign + o.op, label: o.op, info: o.info });
    }
    // Between terms, the two ways to combine them.
    const prior = before.slice(0, before.length - token.length);
    if (!sign && /\S\s+$/.test(prior) && !/(^|\s)OR\s+$/.test(prior)) {
      if ("or".startsWith(word)) options.push({ text: "OR ", label: "OR", info: "match either side" });
      if (!word) options.push({ text: "-", label: "-", info: "exclude the next term" });
    }
  }
  const sections = [];
  if (options.length) sections.push({ title: "Search options", items: options, token });
  if (!before.trim() && history.length) {
    sections.push({ title: "History", clear: true, items: history.map((q) => ({ query: q, label: q })) });
  }
  return sections;
}

// The menu under the search box, built from Obsidian's own search
// suggestion classes so it looks like the search pane's. Clicking or Enter
// takes an item; arrows move between them; Escape closes it.
class SearchMenu {
  constructor(input, { history, clearHistory, apply, noteName }) {
    this.input = input;
    this.history = history;
    this.noteName = noteName;
    this.clearHistory = clearHistory;
    this.apply = apply;
    this.el = null;
    this.items = [];
    this.selected = -1;
  }

  get isOpen() {
    return !!this.el;
  }

  update() {
    const before = this.input.value.slice(0, this.input.selectionStart ?? this.input.value.length);
    // History only while the whole box is empty, as choosing one replaces it.
    const sections = searchSuggestions(before, this.input.value.trim() ? [] : this.history(), this.noteName());
    if (!sections.length) return this.close();
    if (!this.el) {
      // The box's own window, which may be a popout.
      this.el = (this.input.doc || document).body.createDiv({ cls: "suggestion-container mod-search-suggestion sb-todo-search-menu" });
      // Keeps the focus in the search box while an item is clicked.
      this.el.addEventListener("mousedown", (evt) => evt.preventDefault());
    }
    this.el.empty();
    const list = this.el.createDiv({ cls: "suggestion" });
    this.items = [];
    for (const section of sections) {
      const group = list.createDiv({ cls: "suggestion-item mod-complex search-suggest-item mod-group" });
      group.createDiv({ cls: "suggestion-content" })
        .createDiv({ cls: "suggestion-title list-item-part mod-extended" })
        .createSpan({ text: section.title });
      const aux = group.createDiv({ cls: "suggestion-aux" });
      if (section.clear) {
        const clear = aux.createDiv({
          cls: "list-item-part search-suggest-icon clickable-icon",
          attr: { "aria-label": "Clear search history" },
        });
        setIcon(clear, "x");
        clear.addEventListener("click", () => {
          this.clearHistory();
          this.update();
        });
      }
      for (const item of section.items) {
        const el = list.createDiv({ cls: "suggestion-item mod-complex search-suggest-item" });
        const title = el.createDiv({ cls: "suggestion-content" }).createDiv({ cls: "suggestion-title" });
        title.createSpan({ text: item.label });
        if (item.info) title.createSpan({ cls: "search-suggest-info-text", text: item.info });
        el.createDiv({ cls: "suggestion-aux" });
        const index = this.items.length;
        this.items.push({ ...item, token: section.token, el });
        el.addEventListener("mousemove", () => this.select(index));
        el.addEventListener("click", () => this.choose(index));
      }
    }
    // Nothing is picked until the arrows pick it, so Enter keeps the search.
    this.selected = -1;
    this.position();
  }

  position() {
    const rect = this.input.getBoundingClientRect();
    const width = Math.max(rect.width, 260);
    const left = Math.max(8, Math.min(rect.left, (this.input.win || window).innerWidth - width - 8));
    this.el.style.width = `${width}px`;
    this.el.style.left = `${left}px`;
    this.el.style.top = `${rect.bottom + 4}px`;
  }

  select(index) {
    if (!this.items.length) return;
    this.selected = index < 0 ? this.items.length - 1 : index % this.items.length;
    this.items.forEach((item, i) => item.el.toggleClass("is-selected", i === this.selected));
    this.items[this.selected].el.scrollIntoView({ block: "nearest" });
  }

  choose(index) {
    const item = this.items[index];
    if (!item) return;
    const input = this.input;
    if (item.query !== undefined) {
      input.value = item.query;
      input.setSelectionRange(item.query.length, item.query.length);
    } else {
      // Replaces the word being typed, keeping whatever follows the cursor.
      const end = input.selectionStart ?? input.value.length;
      const start = end - item.token.length;
      input.value = input.value.slice(0, start) + item.text + input.value.slice(end);
      const cursor = start + item.text.length;
      input.setSelectionRange(cursor, cursor);
    }
    this.apply(item.query !== undefined);
    this.update();
  }

  // The search box's keys; true when the menu used the key.
  onKeyDown(evt) {
    if (!this.el) return false;
    if (evt.key === "ArrowDown" || evt.key === "ArrowUp") {
      this.select(this.selected + (evt.key === "ArrowDown" ? 1 : -1));
    } else if (evt.key === "Enter" && this.selected >= 0) {
      this.choose(this.selected);
    } else if (evt.key === "Escape") {
      this.close();
    } else {
      return false;
    }
    evt.preventDefault();
    return true;
  }

  close() {
    if (this.el) this.el.remove();
    this.el = null;
    this.items = [];
    this.selected = -1;
  }
}

class TodoPanelSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  // Kept short on purpose: the account, and the line a dropped task becomes.
  // Everything most people never change sits under Advanced, collapsed. The
  // app's client ID and tenant are up front only until signed in.
  display() {
    const { containerEl } = this;
    const s = this.plugin.settings;
    const auth = this.plugin.auth;
    containerEl.empty();

    const account = new Setting(containerEl);
    if (auth.signedIn) {
      account
        .setName(auth.account ? `Signed in as ${auth.account}` : "Signed in to Microsoft")
        .setDesc("The sign-in is kept in Obsidian's secret storage on this device, never in the vault.")
        .addButton((button) => button.setButtonText("Sign out").onClick(() => this.plugin.signOut()));
    } else {
      account
        .setName("Not signed in")
        .setDesc("Register an app in Microsoft Entra (see the README), enter its IDs below, then sign in.")
        .addButton((button) =>
          button
            .setButtonText("Sign in")
            .setCta()
            .onClick(() => this.plugin.signIn())
        );
      this.appSettings(containerEl);
    }

    this.text(
      containerEl,
      "Task template",
      "What an added task becomes. Placeholders: {{title}}, {{due}}, {{date}}, {{list}}, {{link}}, {{id}}.",
      "template"
    );

    const advanced = containerEl.createEl("details", { cls: "sb-todo-advanced" });
    advanced.createEl("summary", { text: "Advanced" });
    if (auth.signedIn) this.appSettings(advanced);
    this.text(advanced, "Due date format", "What {{due}} becomes when a task has a due date. {{date}} is YYYY-MM-DD.", "dueFormat");
    new Setting(advanced)
      .setName("Completed tasks to load (days)")
      .setDesc("How far back the show-completed toggle reaches. 0 turns it off.")
      .addText((text) => {
        text.setValue(String(s.completedDays));
        text.inputEl.addEventListener("change", async () => {
          const days = parseInt(text.getValue(), 10);
          if (!Number.isFinite(days) || days < 0 || days === s.completedDays) return;
          s.completedDays = days;
          await this.plugin.saveSettings();
          this.plugin.panels().forEach((view) => view.refresh());
        });
      });
    new Setting(advanced)
      .setName("Daily note date format")
      .setDesc("How daily note file names spell their date, in Moment.js format. By date groups tasks around an open daily note's date.")
      .addText((text) =>
        text
          .setPlaceholder("YYYY-MM-DD")
          .setValue(s.dailyNoteFormat)
          .onChange(async (value) => {
            s.dailyNoteFormat = value.trim() || "YYYY-MM-DD";
            await this.plugin.saveSettings();
            this.plugin.panels().forEach((view) => view.updateNoteMatches());
          })
      );
    this.authText(
      advanced,
      "Permissions",
      "Microsoft Graph permissions to ask for. Tasks.Read is enough; change it only if your organisation granted the app a different one.",
      "scopes",
      "Tasks.Read"
    );
  }

  // The app registration's IDs, which sign-in needs.
  appSettings(el) {
    this.authText(
      el,
      "Client ID",
      "The Application (client) ID of the app you registered in Microsoft Entra.",
      "clientId",
      "00000000-0000-0000-0000-000000000000"
    );
    this.authText(
      el,
      "Tenant",
      "Your directory (tenant) ID if the app accepts only accounts in your organisation; otherwise common or consumers.",
      "tenant",
      "common"
    );
  }

  // A setting that decides which app or directory to sign in with. It applies
  // when the field is left, not per keystroke, and ends a sign-in that belongs
  // to the old value.
  authText(el, name, desc, key, placeholder) {
    const s = this.plugin.settings;
    const auth = this.plugin.auth;
    new Setting(el)
      .setName(name)
      .setDesc(desc)
      .addText((text) => {
        text.setPlaceholder(placeholder).setValue(s[key]);
        text.inputEl.addEventListener("change", async () => {
          const value = text.getValue().trim();
          if (value === s[key]) return;
          const wasSignedIn = auth.signedIn;
          s[key] = value;
          // The sign-in dialog's remembered IDs would otherwise undo this.
          if (this.plugin.signInDraft && key in this.plugin.signInDraft) this.plugin.signInDraft[key] = value;
          await this.plugin.saveSettings();
          if (wasSignedIn && !auth.matchesSettings()) this.plugin.signOut();
        });
      });
  }

  text(el, name, desc, key) {
    new Setting(el)
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

// A list item, with its checkbox mark (if any) and its text.
const LIST_ITEM = /^(?:[-*+]|\d+[.)])\s+(?:\[(.)\]\s+)?(.*)$/;

// The task-like lines of a note: each raw trimmed line, and each list item's
// text reduced by normalizeText, mapped to whether any such line is checked.
// Only [x] or [X] counts as checked; [<], [>], [-] and other marks do not.
// Each also records the first line it is on, for jumping to it.
function noteTaskLines(text) {
  const raw = new Map();
  const titles = new Map();
  (text || "").split(/\r?\n/).forEach((line, index) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    if (!raw.has(trimmed)) raw.set(trimmed, index);
    const item = trimmed.match(LIST_ITEM);
    if (!item) return;
    const key = normalizeText(item[2]);
    const checked = item[1] === "x" || item[1] === "X";
    const seen = titles.get(key);
    titles.set(key, { checked: Boolean(seen && seen.checked) || checked, line: seen ? seen.line : index });
  });
  return { raw, titles };
}

// The panel's tasks keyed like note lines, for the links in notes: each task's
// normalized title and its rendered line, mapped to the task's ID. When two
// tasks share a title, an open one wins over a completed one.
function buildTaskIndex(lists, render) {
  const titles = new Map();
  const raw = new Map();
  const completed = new Set();
  for (const list of lists) {
    for (const task of list.tasks) {
      const title = normalizedTitle(task);
      if (title && (!titles.has(title) || (completed.has(titles.get(title)) && !task.completed))) titles.set(title, task.id);
      const line = render(task, list);
      if (line && (!raw.has(line) || (completed.has(raw.get(line)) && !task.completed))) raw.set(line, task.id);
      if (task.completed) completed.add(task.id);
    }
  }
  return titles.size || raw.size ? { titles, raw } : null;
}

// Longer lines are never task lines, and are not worth normalizing.
const MAX_TASK_LINE = 1000;

// The ID of the task a note line holds, or null: the same matching as the
// "in note" markers, for one line.
function taskIdForLine(index, text) {
  if (!text || text.length > MAX_TASK_LINE) return null;
  const trimmed = text.trim();
  if (!trimmed) return null;
  const item = trimmed.match(LIST_ITEM);
  const id = item && index.titles.get(normalizeText(item[2]));
  return id || index.raw.get(trimmed) || null;
}

// The editor side of the links: a small button at the end of each line that
// holds a panel task. Only the lines on screen are looked at, and only when
// the text, the visible part or the tasks change; moving the cursor costs
// nothing. Each distinct line text is looked up once, so scrolling back and
// typing elsewhere reuse earlier answers.
function taskLinkExtension(plugin) {
  const redrawEffect = StateEffect.define();
  const decorations = new Map(); // line text → decoration or null
  let cachedIndex = null;

  class TaskLinkWidget extends WidgetType {
    constructor(id) {
      super();
      this.id = id;
    }
    eq(other) {
      return other.id === this.id;
    }
    toDOM() {
      const el = document.createElement("span");
      el.className = "sb-todo-reveal";
      el.setAttribute("role", "button");
      el.setAttribute("aria-label", "Show in To Do panel");
      // Keeps the cursor and focus where they are.
      el.addEventListener("mousedown", (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
      });
      el.addEventListener("click", (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        plugin.revealTask(this.id).catch((e) => new Notice(`Could not show the task: ${e.message}`));
      });
      return el;
    }
    // The editor leaves clicks on the button alone.
    ignoreEvent() {
      return true;
    }
  }

  const decorationFor = (index, text) => {
    if (index !== cachedIndex || decorations.size > 2000) {
      decorations.clear();
      cachedIndex = index;
    }
    let deco = decorations.get(text);
    if (deco === undefined) {
      const id = taskIdForLine(index, text);
      deco = id ? Decoration.widget({ widget: new TaskLinkWidget(id), side: 1 }) : null;
      decorations.set(text, deco);
    }
    return deco;
  };

  return ViewPlugin.fromClass(
    class {
      constructor(view) {
        this.view = view;
        this.decorations = this.build();
        plugin.taskLinkViews.add(this);
      }
      update(update) {
        if (update.docChanged || update.viewportChanged || update.transactions.some((tr) => tr.effects.some((e) => e.is(redrawEffect)))) {
          this.decorations = this.build();
        }
      }
      redraw() {
        this.view.dispatch({ effects: redrawEffect.of(null) });
      }
      destroy() {
        plugin.taskLinkViews.delete(this);
      }
      build() {
        const index = plugin.taskIndexNow();
        if (!index) return Decoration.none;
        const doc = this.view.state.doc;
        const builder = new RangeSetBuilder();
        let last = 0;
        for (const { from, to } of this.view.visibleRanges) {
          let line = doc.lineAt(from);
          // Ranges can share a line; each line is decorated once, in order.
          if (line.number <= last) {
            if (last >= doc.lines) break;
            line = doc.line(last + 1);
          }
          while (line.from <= to) {
            const deco = decorationFor(index, line.text);
            if (deco) builder.add(line.to, line.to, deco);
            last = line.number;
            if (line.number === doc.lines) break;
            line = doc.line(line.number + 1);
          }
        }
        return builder.finish();
      }
    },
    { decorations: (links) => links.decorations }
  );
}

// The line of the note that holds this task, or -1. Same matching as the
// "in note" markers.
function findTaskLine(lines, title, rendered) {
  const match = title && lines.titles.get(title);
  if (match) return match.line;
  return lines.raw.has(rendered) ? lines.raw.get(rendered) : -1;
}

// What makes two task texts "the same": case, indentation and runs of
// whitespace do not count, and neither does a tail of emojis, dates and links,
// such as "📅 2026-11-20", "✅ 16/11/2026" or "[To Do](https://…)".
const TRAILING_NOISE = new RegExp(
  "(?:\\s*(?:" +
    "\\[[^\\]]*\\]\\([^)]*\\)" + // Markdown link
    "|\\d{4}-\\d{2}-\\d{2}(?:[ T]\\d{1,2}:\\d{2})?" + // 2026-11-20, optional time
    "|\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}" + // 20/11/2026, 20.11.26
    // One symbol per repetition: a "+" here as well would nest two
    // quantifiers and make long emoji runs backtrack for seconds.
    "|[\\p{Extended_Pictographic}\\u2190-\\u21FF\\u2B00-\\u2BFF\\uFE0F\\u200D]" + // emojis and arrows
    "))+\\s*$",
  "u"
);

// normalizeText of a task's title, computed once per fetched task.
const titleCache = new WeakMap();
function normalizedTitle(task) {
  if (!titleCache.has(task)) titleCache.set(task, normalizeText(task.title));
  return titleCache.get(task);
}

// A short, stable lowercase hex hash (FNV-1a), for IDs derived from names.
function hashString(text) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function normalizeText(text) {
  return (text || "").replace(TRAILING_NOISE, "").replace(/\s+/g, " ").trim().toLowerCase();
}

function fill(template, values) {
  return (template || "").replace(/\{\{\s*(\w+)\s*\}\}/g, (whole, key) =>
    Object.prototype.hasOwnProperty.call(values, key) ? values[key] : whole
  );
}

function localToday() {
  return isoDay(new Date());
}

// A local date as YYYY-MM-DD.
function isoDay(date) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
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
    // Bumped whenever a session ends or a new one begins, so token responses
    // and task fetches started under an older session are thrown away.
    this.epoch = 0;
    this.secretId = `${SESSION_SECRET_PREFIX}-${hashString(plugin.app.vault.getName())}`;
    this.session = this.loadSession();
  }

  get settings() {
    return this.plugin.settings;
  }

  // A stored session counts only for the client ID and tenant now set.
  get signedIn() {
    return Boolean(this.session && this.session.refresh_token && this.matchesSettings());
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

  // The app and directory a request is made for, fixed when it starts.
  context() {
    return { clientId: this.settings.clientId.trim(), tenant: this.tenant(), epoch: this.epoch };
  }

  endpoint(path, tenant = this.tenant()) {
    return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/${path}`;
  }

  scope() {
    const asked = (this.settings.scopes || "").split(/[\s,]+/).filter(Boolean);
    return [...new Set([...asked, "offline_access", "openid", "profile"])].join(" ");
  }

  // ── session storage ──

  loadSession() {
    const store = this.plugin.app.secretStorage;
    try {
      let raw = store.getSecret(this.secretId);
      if (!raw) {
        // Move a session saved by 1.0.3 or earlier under the legacy ID.
        raw = store.getSecret(SESSION_SECRET_PREFIX);
        if (raw) {
          store.setSecret(this.secretId, raw);
          store.setSecret(SESSION_SECRET_PREFIX, "");
        }
      }
      return raw ? JSON.parse(raw) : null;
    } catch (_) {
      return null;
    }
  }

  saveSession(session) {
    this.session = session;
    // Secret storage has no delete; an empty value reads back as signed out.
    this.plugin.app.secretStorage.setSecret(this.secretId, session ? JSON.stringify(session) : "");
  }

  signOut() {
    this.epoch++;
    this.saveSession(null);
  }

  // ── device code flow ──

  async startDeviceFlow() {
    const ctx = this.context();
    const res = await postForm(this.endpoint("devicecode", ctx.tenant), {
      client_id: ctx.clientId,
      scope: this.scope(),
    });
    if (res.status !== 200) throw new Error(describeAuthError(res.json));
    // The token poll must use the same app and directory as the code.
    return { ...res.json, ctx };
  }

  // Polls until the user finishes signing in. `cancel` is an object whose
  // `cancelled` flag stops the loop, for a closed dialog or an unloaded plugin.
  async completeDeviceFlow(flow, cancel) {
    const ctx = flow.ctx || this.context();
    let interval = Math.max(1, Number(flow.interval) || 5);
    const deadline = Date.now() + (Number(flow.expires_in) || 900) * 1000;
    while (!cancel.cancelled) {
      await sleep(interval * 1000);
      if (cancel.cancelled) break;
      if (Date.now() > deadline) throw new Error("The sign-in code expired. Start again.");
      const res = await postForm(this.endpoint("token", ctx.tenant), {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: ctx.clientId,
        device_code: flow.device_code,
      });
      // Closed or unloaded while that request was out: stay signed out.
      if (cancel.cancelled) break;
      if (res.status === 200) {
        // A new session: anything started under the old one is now stale.
        this.epoch++;
        this.storeTokens(res.json, ctx);
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

  storeTokens(tokens, ctx) {
    const previous = this.session || {};
    const claims = jwtClaims(tokens.id_token);
    this.saveSession({
      client_id: ctx.clientId,
      tenant: ctx.tenant,
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
    if (!this.signedIn) throw new NotSignedInError("Not signed in to Microsoft.");
    return this.session.access_token;
  }

  async refresh() {
    const ctx = { clientId: this.session.client_id, tenant: this.session.tenant, epoch: this.epoch };
    const res = await postForm(this.endpoint("token", ctx.tenant), {
      grant_type: "refresh_token",
      client_id: ctx.clientId,
      refresh_token: this.session.refresh_token,
      scope: this.scope(),
    });
    // Signed out while that request was out: do not bring the session back.
    if (ctx.epoch !== this.epoch) throw new NotSignedInError("Not signed in to Microsoft.");
    if (res.status === 200) {
      this.storeTokens(res.json, ctx);
      return;
    }
    const error = res.json && res.json.error;
    if (error === "invalid_grant" || error === "interaction_required" || error === "invalid_client") {
      this.signOut();
      const expired = new NotSignedInError(`Your Microsoft sign-in has expired. Sign in again. (${describeAuthError(res.json)})`);
      expired.expired = true;
      throw expired;
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
    // Signing in means leaving for a browser and coming back. A click on
    // the dimmed background on return would close the dialog and lose what
    // was typed, so only Escape and the close button close it.
    const bg = this.containerEl.querySelector(".modal-bg");
    if (bg) bg.addEventListener("click", (evt) => evt.stopImmediatePropagation(), true);
    this.askForApp();
  }

  // The app registration's IDs, filled in with what was entered last, so
  // the order they are typed in never matters and a closed dialog loses
  // nothing.
  askForApp() {
    const { contentEl } = this;
    const s = this.plugin.settings;
    const draft = this.plugin.signInDraft || (this.plugin.signInDraft = { clientId: s.clientId.trim(), tenant: this.plugin.auth.tenant() });
    let started = false;
    contentEl.createDiv({
      cls: "sb-todo-signin-status",
      text: "Enter the IDs of the app you registered in Microsoft Entra. The plugin's README explains how to register one.",
    });
    new Setting(contentEl)
      .setName("Client ID")
      .setDesc("Application (client) ID")
      .addText((text) =>
        text
          .setPlaceholder("00000000-0000-0000-0000-000000000000")
          .setValue(draft.clientId)
          .onChange((v) => (draft.clientId = v.trim()))
      );
    new Setting(contentEl)
      .setName("Tenant")
      .setDesc("Directory (tenant) ID for an app limited to your organisation; otherwise common or consumers.")
      .addText((text) => text.setPlaceholder("common").setValue(draft.tenant).onChange((v) => (draft.tenant = v.trim())));
    new Setting(contentEl).addButton((button) =>
      button
        .setButtonText("Continue")
        .setCta()
        .onClick(async () => {
          if (started) return;
          if (!draft.clientId) {
            new Notice("Enter the client ID.");
            return;
          }
          started = true;
          s.clientId = draft.clientId;
          s.tenant = draft.tenant || "common";
          try {
            await this.plugin.saveSettings();
          } catch (e) {
            started = false;
            new Notice(`Could not save the IDs: ${e.message}`);
            return;
          }
          this.plugin.refreshSettingTab();
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
      if (!this.cancel.cancelled) this.failed(status, `Could not start sign-in: ${e.message}`);
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
      this.plugin.signInDraft = null;
      this.close();
      this.plugin.signedIn();
    } catch (e) {
      if (this.cancel.cancelled) return;
      codeRow.remove();
      actions.remove();
      this.failed(status, e.message);
    }
  }

  // A failed sign-in can go back to the IDs, which are often what was wrong.
  failed(status, message) {
    status.setText(message);
    status.addClass("sb-todo-error");
    new Setting(this.contentEl).addButton((button) =>
      button.setButtonText("Back").onClick(() => {
        this.contentEl.empty();
        this.askForApp();
      })
    );
  }

  onClose() {
    this.cancel.cancelled = true;
    this.contentEl.empty();
    if (this.plugin.signInModal === this) this.plugin.signInModal = null;
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
  // To Do sends a completion day as that bare date at midnight, so compare
  // with the local calendar date, not a UTC instant.
  const since = new Date();
  since.setDate(since.getDate() - completedDays);
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${since.getFullYear()}-${pad(since.getMonth() + 1)}-${pad(since.getDate())}T00:00:00`;
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
    notes: body.contentType === "html" ? htmlToText(body.content) : (body.content || "").trim(),
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

// Adds text as a line of its own at a line (the cursor's by default): on that
// line when it is blank, otherwise on a new line below it, with the same
// indentation.
function insertOnOwnLine(editor, text, line = editor.getCursor().line) {
  const current = editor.getLine(line);
  const indent = current.match(/^[ \t]*/)[0];
  // A line with only a list marker ("- ", "- [ ] ") is taken over too.
  const blank = !current.trim() || isBareListMarker(current);
  const target = blank ? line : line + 1;
  editor.transaction({
    changes: [
      blank
        ? { from: { line, ch: 0 }, to: { line, ch: current.length }, text: indent + text }
        : { from: { line, ch: current.length }, to: { line, ch: current.length }, text: `\n${indent}${text}` },
    ],
  });
  editor.setCursor({ line: target, ch: indent.length + text.length });
}

function isBareListMarker(line) {
  return /^\s*(?:[-*+]|\d+[.)])\s*(?:\[.\]\s*)?$/.test(line);
}

// Writes text into an element with its links clickable: web links open in
// the browser, obsidian:// links open their note. A click on a link only
// follows it, so it never also inserts or shows the task, and dragging from a
// link drags the task, not the link.
function appendLinkedText(el, text, app) {
  for (const part of splitLinks(text)) {
    if (!part.url) {
      el.appendText(part.text);
      continue;
    }
    const a = el.createEl("a", { cls: "external-link", text: part.text, href: part.url, attr: { target: "_blank", rel: "noopener", draggable: "false" } });
    a.addEventListener("click", (evt) => {
      evt.stopPropagation();
      if (!part.url.startsWith("obsidian://")) return;
      // A note in this vault opens here, like a wiki link. So does one that
      // names another vault but exists here too: a vault synced between
      // devices can have a different name on each. Anything else goes to
      // Obsidian's own handler for obsidian:// links.
      evt.preventDefault();
      const link = obsidianLink(part.url);
      const here = link && (link.vault === null || link.vault === app.vault.getName() || app.metadataCache.getFirstLinkpathDest(link.file, ""));
      if (here) app.workspace.openLinkText(link.file, "", evt.ctrlKey || evt.metaKey);
      else window.open(part.url);
    });
  }
}

// The vault (null when not named) and note an obsidian://open link opens;
// null for any other link.
function obsidianLink(url) {
  // Read by hand: Chromium does not parse hosts of custom schemes the same
  // way in every version.
  const match = /^obsidian:\/\/open\/?\?(.*)$/i.exec(url);
  if (!match) return null;
  const params = new URLSearchParams(match[1]);
  const vault = params.get("vault");
  const file = params.get("file");
  return file ? { vault, file } : null;
}

// Splits plain text into text and link parts (web and obsidian:// links), for
// showing links as links. Punctuation that ends a sentence is not part of the
// link, and neither is a closing bracket the link did not open.
function splitLinks(text) {
  const parts = [];
  const pattern = /(?:https?|obsidian):\/\/[^\s<>"]+/g;
  let last = 0;
  let match;
  while ((match = pattern.exec(text || "")) !== null) {
    let url = match[0];
    while (/[.,;:!?'\]]$/.test(url) || (url.endsWith(")") && !url.includes("("))) url = url.slice(0, -1);
    if (match.index > last) parts.push({ text: text.slice(last, match.index) });
    parts.push({ text: url, url });
    last = match.index + url.length;
    pattern.lastIndex = last;
  }
  if (last < (text || "").length) parts.push({ text: text.slice(last) });
  return parts;
}

// After a drop, the editor selects the text it inserted. If that text landed
// inside a line with other text, take it out again and put it on a new line
// below that line, or above it when dropped at the line's start, so no line
// is split. The new line gets the same indentation as the line dropped on,
// so a task dropped on a nested list item lands beside it. Does nothing when
// the selection is not the dropped text.
function placeOnOwnLine(editor, text) {
  if (!text || text.includes("\n")) return;
  const selection = editor.listSelections()[0];
  if (!selection) return;
  const [from, to] = [selection.anchor, selection.head].sort((a, b) => a.line - b.line || a.ch - b.ch);
  if (from.line !== to.line || editor.getRange(from, to) !== text) return;
  const line = editor.getLine(from.line);
  const before = line.slice(0, from.ch);
  const after = line.slice(to.ch);
  if (!before.trim() && !after.trim()) return;
  const remaining = before + after;
  const indent = remaining.match(/^[ \t]*/)[0];
  // Dropped on an empty bullet ("- ", "- [ ] "): the task replaces it.
  if (isBareListMarker(remaining)) {
    editor.transaction({ changes: [{ from: { line: from.line, ch: 0 }, to: { line: from.line, ch: line.length }, text: indent + text }] });
    editor.setSelection({ line: from.line, ch: indent.length }, { line: from.line, ch: indent.length + text.length });
    return;
  }
  const atStart = !before.trim();
  const target = atStart ? from.line : from.line + 1;
  // One transaction, so a single undo reverts the move.
  editor.transaction({
    changes: [
      {
        from: { line: from.line, ch: 0 },
        to: { line: from.line, ch: line.length },
        text: atStart ? `${indent}${text}\n${remaining}` : `${remaining}\n${indent}${text}`,
      },
    ],
  });
  editor.setSelection({ line: target, ch: indent.length }, { line: target, ch: indent.length + text.length });
}

// Plain text from a task's HTML notes, for display as text only (never as
// HTML): line breaks kept, tags dropped, common entities decoded.
function htmlToText(html) {
  return String(html || "")
    // A link keeps its address: "text (https://…)", or the address alone.
    .replace(/<a\b[^>]*\bhref\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi, (whole, href, inner) => {
      const label = inner.replace(/<[^>]*>/g, "").trim();
      return !label || label === href ? href : `${label} (${href})`;
    })
    .replace(/<(br|\/p|\/div|\/li|\/h[1-6])\b[^>]*>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function postForm(url, fields) {
  const res = await requestUrl({
    url,
    method: "POST",
    contentType: "application/x-www-form-urlencoded",
    body: new URLSearchParams(fields).toString(),
    throw: false,
  });
  return { status: res.status, headers: res.headers, json: safeJson(res) };
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
