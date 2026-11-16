// Tests for sign-in, Microsoft Graph reading and note matching.
// Run with: node test/run.js   (TZ is pinned so due-date checks are stable)
process.env.TZ = "Europe/Copenhagen";

const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");

// Load main.js against the stub, exposing its internal functions.
const stubPath = path.join(__dirname, "obsidian-stub.js");
const realResolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  return request === "obsidian" ? stubPath : realResolve.call(this, request, ...rest);
};
const mainPath = path.join(__dirname, "..", "src", "plugin.js");
const source =
  fs.readFileSync(mainPath, "utf8") +
  "\nmodule.exports.__internals = { MicrosoftAuth, noteTaskLines, normalizeText, describeAuthError, taskFilter, placeOnOwnLine, splitLinks };\n";
const mod = new Module(mainPath, module);
mod.filename = mainPath;
mod.paths = Module._nodeModulePaths(path.dirname(mainPath));
mod._compile(source, mainPath);
const TodoPanelPlugin = mod.exports;
const { MicrosoftAuth, noteTaskLines, normalizeText, describeAuthError, taskFilter, placeOnOwnLine, splitLinks } = TodoPanelPlugin.__internals;

global.window = { setTimeout: (fn) => setImmediate(fn), clearTimeout() {}, setInterval, clearInterval };
global.atob = global.atob || ((s) => Buffer.from(s, "base64").toString("binary"));

const LEGACY_SESSION = "todo-panel-microsoft-session";
const idToken = `h.${Buffer.from(JSON.stringify({ preferred_username: "user@example.com" })).toString("base64url")}.s`;

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

async function signedInPlugin() {
  global.__requests = [];
  global.__data = { clientId: "client", tenant: "tenant" };
  const plugin = new TodoPanelPlugin();
  await plugin.onload();
  global.__handler = () => ({ body: { access_token: "AT1", refresh_token: "RT1", expires_in: 3600, id_token: idToken } });
  await plugin.auth.completeDeviceFlow({ device_code: "d", interval: 1 }, { cancelled: false });
  global.__requests = [];
  return plugin;
}

test("drops settings from the external-command versions", async () => {
  global.__data = { command: "old", timeoutSeconds: 5 };
  const plugin = new TodoPanelPlugin();
  await plugin.onload();
  assert.strictEqual(plugin.settings.command, undefined);
  assert.strictEqual(plugin.settings.timeoutSeconds, undefined);
});

test("device code flow waits through pending and slow_down", async () => {
  global.__requests = [];
  global.__data = { clientId: "client", tenant: "tenant" };
  const plugin = new TodoPanelPlugin();
  await plugin.onload();
  let polls = 0;
  global.__handler = (req) => {
    if (req.url.endsWith("/devicecode")) {
      return { body: { user_code: "ABC", device_code: "dev", verification_uri: "https://microsoft.com/devicelogin", interval: 1 } };
    }
    polls++;
    if (polls === 1) return { status: 400, body: { error: "authorization_pending" } };
    if (polls === 2) return { status: 400, body: { error: "slow_down" } };
    return { body: { access_token: "AT", refresh_token: "RT", expires_in: 3600, id_token: idToken } };
  };
  const flow = await plugin.auth.startDeviceFlow();
  assert.ok(global.__requests[0].url.startsWith("https://login.microsoftonline.com/tenant/oauth2/v2.0/devicecode"));
  assert.ok(global.__requests[0].body.includes("scope=Tasks.Read+offline_access+openid+profile"));
  await plugin.auth.completeDeviceFlow(flow, { cancelled: false });
  assert.strictEqual(polls, 3);
  assert.ok(plugin.auth.signedIn);
  assert.strictEqual(plugin.auth.account, "user@example.com");
});

test("keeps tokens in secret storage, never in plugin data", async () => {
  const plugin = await signedInPlugin();
  await plugin.saveSettings();
  assert.ok(!JSON.stringify(global.__saved).includes("RT1"));
  assert.ok(plugin.app.secrets[plugin.auth.secretId].includes("RT1"));
});

test("a declined or cancelled sign-in stops", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ status: 400, body: { error: "authorization_declined" } });
  await assert.rejects(plugin.auth.completeDeviceFlow({ device_code: "d", interval: 1 }, { cancelled: false }), /declined/);
  await assert.rejects(plugin.auth.completeDeviceFlow({ device_code: "d", interval: 1 }, { cancelled: true }), /cancelled/);
});

test("concurrent callers share one refresh and keep the rotated token", async () => {
  const plugin = await signedInPlugin();
  plugin.auth.session.expires_at = 0;
  let refreshes = 0;
  global.__handler = (req) => {
    refreshes++;
    assert.ok(req.body.includes("grant_type=refresh_token") && req.body.includes("refresh_token=RT1"));
    return { body: { access_token: "AT2", refresh_token: "RT2", expires_in: 3600 } };
  };
  const [a, b] = await Promise.all([plugin.auth.accessToken(), plugin.auth.accessToken()]);
  assert.strictEqual(refreshes, 1);
  assert.strictEqual(a, "AT2");
  assert.strictEqual(b, "AT2");
  assert.strictEqual(plugin.auth.session.refresh_token, "RT2");
  assert.strictEqual(plugin.auth.account, "user@example.com");
});

test("an expired refresh token signs out with a short message", async () => {
  const plugin = await signedInPlugin();
  plugin.auth.session.expires_at = 0;
  global.__handler = () => ({
    status: 400,
    body: { error: "invalid_grant", error_description: "AADSTS70008: The refresh token has expired.\r\nTrace ID: x" },
  });
  await assert.rejects(plugin.auth.accessToken(), /Sign in again\. \(AADSTS70008: The refresh token has expired\.\)/);
  assert.strictEqual(plugin.auth.signedIn, false);
});

test("a session issued for another app is dropped", async () => {
  const plugin = await signedInPlugin();
  const stored = JSON.stringify({ client_id: "other", tenant: "tenant", refresh_token: "r" });
  plugin.app.secrets[plugin.auth.secretId] = stored;
  assert.strictEqual(new MicrosoftAuth(plugin).signedIn, false);
  assert.strictEqual(plugin.app.secrets[plugin.auth.secretId], stored, "not wiped, in case another vault uses it");
});

test("reads lists and tasks: 401 retry, 429 back-off, paging, mapping, sorting", async () => {
  const plugin = await signedInPlugin();
  let listCalls = 0;
  let taskCalls = 0;
  global.__handler = (req) => {
    if (req.url.endsWith("/token")) return { body: { access_token: "AT3", refresh_token: "RT3", expires_in: 3600 } };
    if (req.url.endsWith("/me/todo/lists")) {
      listCalls++;
      if (listCalls === 1) return { status: 401, body: { error: { message: "expired" } } };
      return { body: { value: [{ id: "L1", displayName: "Inbox" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?page=2" } };
    }
    if (req.url.includes("lists?page=2")) return { body: { value: [{ id: "L2", displayName: "Empty" }] } };
    if (req.url.includes("/lists/L1/tasks")) {
      if (taskCalls++ === 0) return { status: 429, headers: { "Retry-After": "1" } };
      return {
        body: {
          value: [
            { id: "t1", title: "Open undated", status: "notStarted" },
            {
              id: "t2",
              title: "Due",
              status: "notStarted",
              importance: "high",
              dueDateTime: { dateTime: "2026-11-19T23:00:00.0000000", timeZone: "UTC" },
              body: { contentType: "text", content: " a note " },
            },
            { id: "t3", title: "Done", status: "completed", completedDateTime: { dateTime: "2026-11-16T00:00:00.0000000", timeZone: "UTC" } },
          ],
        },
      };
    }
    if (req.url.includes("/lists/L2/tasks")) return { body: { value: [] } };
    throw new Error(`Unexpected request ${req.url}`);
  };
  const lists = await plugin.fetchLists();
  assert.strictEqual(listCalls, 2, "retried once after 401");
  assert.strictEqual(taskCalls, 2, "retried once after 429");
  assert.deepStrictEqual(lists.map((l) => l.name), ["Inbox", "Empty"]);
  const tasks = lists[0].tasks;
  assert.deepStrictEqual(tasks.map((t) => t.title), ["Due", "Open undated", "Done"]);
  assert.strictEqual(tasks[0].due_date, "2026-11-20", "due day is the local day");
  assert.strictEqual(tasks[0].notes, "a note");
  assert.strictEqual(tasks[2].completed_date, "2026-11-16", "completion day taken as is");
  const taskRequest = global.__requests.find((r) => r.url.includes("/lists/L1/tasks"));
  assert.ok(decodeURIComponent(taskRequest.url).includes("status ne 'completed' or completedDateTime/dateTime ge '"));
  assert.strictEqual(taskRequest.headers.Authorization, "Bearer AT3");
});

test("the panel keeps the lists it fetched", async () => {
  const plugin = await signedInPlugin();
  const lists = [{ id: "L1", name: "Inbox", tasks: [{ id: "t", title: "Task", due_date: null }] }];
  plugin.fetchLists = async () => lists;
  const view = global.__viewFactory({});
  await view.refresh();
  assert.strictEqual(view.error, null);
  assert.deepStrictEqual(view.lists, lists);
  assert.strictEqual(view.groups().length, 1);
});

test("opening the panel does not wait for tasks to load", async () => {
  const plugin = await signedInPlugin();
  let finishFetch;
  plugin.fetchLists = () => new Promise((resolve) => (finishFetch = resolve));
  const view = global.__viewFactory({});
  // A minimal element API, enough for onOpen to build its toolbar.
  const el = () => {
    const node = {
      empty() {},
      addClass() {},
      toggleClass() {},
      setAttr() {},
      setText() {},
      addEventListener() {},
      querySelectorAll: () => [],
      createDiv: () => el(),
      createEl: () => el(),
      createSpan: () => el(),
    };
    return node;
  };
  view.app = plugin.app;
  view.contentEl = el();
  global.__layoutReady = null;
  const opened = await Promise.race([view.onOpen().then(() => "opened"), new Promise((r) => setTimeout(() => r("blocked"), 50))]);
  assert.strictEqual(opened, "opened", "onOpen returned before tasks loaded");
  assert.ok(global.__layoutReady, "fetch deferred until the layout is ready");
  global.__layoutReady();
  await new Promise((resolve) => setImmediate(resolve)); // the fetch waits for an idle moment
  assert.ok(view.loading, "fetch running in the background");
  finishFetch([]);
});

test("an expired sign-in found while loading tasks is shown in the panel", async () => {
  const plugin = await signedInPlugin();
  plugin.saveListCache([{ id: "L1", name: "Inbox", tasks: [] }]);
  plugin.auth.session.expires_at = 0;
  global.__handler = () => ({ status: 400, body: { error: "invalid_grant", error_description: "AADSTS70008: expired." } });
  const view = global.__viewFactory({});
  await view.refresh();
  assert.match(view.error || "", /sign-in has expired/);
  assert.deepStrictEqual(view.lists, []);
  assert.strictEqual(plugin.app.loadLocalStorage("microsoft-todo-panel-lists"), null);
});

test("a session under the legacy secret ID moves to this vault's ID", async () => {
  const plugin = await signedInPlugin();
  const session = plugin.app.secrets[plugin.auth.secretId];
  plugin.app.secrets[plugin.auth.secretId] = "";
  plugin.app.secrets[LEGACY_SESSION] = session;
  const auth = new MicrosoftAuth(plugin);
  assert.ok(auth.signedIn);
  assert.strictEqual(plugin.app.secrets[auth.secretId], session);
  assert.strictEqual(plugin.app.secrets[LEGACY_SESSION], "");
  assert.match(auth.secretId, /^[a-z0-9-]+$/);
});

test("signing out while a token refresh is in flight stays signed out", async () => {
  const plugin = await signedInPlugin();
  plugin.auth.session.expires_at = 0;
  let answer;
  global.__handler = () => ({ body: { access_token: "AT2", refresh_token: "RT2", expires_in: 3600 } });
  const pending = plugin.auth.accessToken();
  plugin.auth.signOut();
  await assert.rejects(pending, /Not signed in/);
  assert.strictEqual(plugin.auth.signedIn, false);
  assert.strictEqual(plugin.app.secrets[plugin.auth.secretId], "");
});

test("closing the sign-in dialog while the token request is out stays signed out", async () => {
  global.__requests = [];
  global.__data = { clientId: "client", tenant: "tenant" };
  const plugin = new TodoPanelPlugin();
  await plugin.onload();
  const cancel = { cancelled: false };
  global.__handler = () => {
    cancel.cancelled = true; // the user closes the dialog while this is out
    return { body: { access_token: "AT", refresh_token: "RT", expires_in: 3600 } };
  };
  await assert.rejects(plugin.auth.completeDeviceFlow({ device_code: "d", interval: 1 }, cancel), /cancelled/);
  assert.strictEqual(plugin.auth.signedIn, false);
});

test("lists fetched for a session that ended are thrown away", async () => {
  const plugin = await signedInPlugin();
  let finish;
  plugin.fetchLists = () => new Promise((resolve) => (finish = resolve));
  const view = global.__viewFactory({});
  const loading = view.refresh();
  plugin.auth.signOut();
  finish([{ id: "L1", name: "Old account", tasks: [] }]);
  await loading;
  assert.deepStrictEqual(view.lists, []);
  assert.strictEqual(plugin.app.loadLocalStorage("microsoft-todo-panel-lists"), null);
});

test("a refresh asked for while one runs is queued, not dropped", async () => {
  const plugin = await signedInPlugin();
  let calls = 0;
  const finishers = [];
  plugin.fetchLists = () => {
    calls++;
    return new Promise((resolve) => finishers.push(resolve));
  };
  const view = global.__viewFactory({});
  const first = view.refresh();
  view.refresh();
  finishers[0]([]);
  await first;
  assert.strictEqual(calls, 2);
  finishers[1]([]);
});

test("the last lists are cached for this account and shown at once", async () => {
  const plugin = await signedInPlugin();
  const lists = [{ id: "L1", name: "Inbox", tasks: [] }];
  plugin.fetchLists = async () => lists;
  await global.__viewFactory({}).refresh();
  assert.deepStrictEqual(plugin.loadListCache().lists, lists);
  plugin.auth.session.account = "someone-else@example.com";
  assert.strictEqual(plugin.loadListCache(), null, "another account never sees them");
});

test("long emoji runs in a note do not hang the matcher", () => {
  const family = "\u{1F468}\u200D\u{1F469}\u200D\u{1F467}\u200D\u{1F466}";
  const started = Date.now();
  normalizeText(`- [ ] Party ${family.repeat(40)} x`);
  normalizeText(`Plan ${"\u{1F389}".repeat(2000)} x`);
  assert.ok(Date.now() - started < 1000, `took ${Date.now() - started} ms`);
  assert.strictEqual(normalizeText(`Party \u{1F389}\u{1F389} 📅 2026-11-20`), "party");
});

test("the completed filter starts at local midnight N days ago", () => {
  const d = new Date();
  d.setDate(d.getDate() - 7);
  const pad = (n) => String(n).padStart(2, "0");
  const expected = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T00:00:00`;
  assert.ok(taskFilter(7).endsWith(`ge '${expected}'`));
  assert.strictEqual(taskFilter(0), "status ne 'completed'");
});

test("clicking a task the note holds selects its line; others are added below the cursor", async () => {
  const plugin = await signedInPlugin();
  let text = ["# Day", "", "- [ ] Send slides 📅 2026-11-20", "    - [x] Book venue"];
  const calls = [];
  const editor = {
    getValue: () => text.join("\n"),
    getCursor: () => ({ line: 3, ch: 5 }),
    setCursor: (pos) => calls.push(["cursor", pos]),
    transaction({ changes }) {
      const c = changes[0];
      const lines = text.slice();
      lines[c.from.line] = lines[c.from.line].slice(0, c.from.ch) + c.text + lines[c.from.line].slice(c.to.ch);
      text = lines.join("\n").split("\n");
    },
    getLine: (n) => text[n],
    setSelection: (from, to) => calls.push(["select", from, to]),
    scrollIntoView: () => calls.push(["scroll"]),
    focus: () => calls.push(["focus"]),
  };
  // Focus is in the panel, so there is no active Markdown view: the note is
  // the one that was active last in the main area, here in reading view.
  const { MarkdownView } = require("./obsidian-stub.js");
  const view = Object.assign(new MarkdownView(), {
    editor,
    leaf: { id: "note" },
    mode: "preview",
    getViewData: () => text.join("\n"),
    getMode() {
      return this.mode;
    },
    getState: () => ({ file: "Day.md" }),
    async setState(state) {
      this.mode = state.mode;
    },
  });
  plugin.app.workspace.getActiveViewOfType = () => null;
  plugin.app.workspace.getMostRecentLeaf = () => ({ view });
  plugin.app.workspace.setActiveLeaf = (leaf) => calls.push(["activate", leaf.id]);
  const list = { id: "L", name: "Inbox" };
  await plugin.showOrInsert({ id: "1", title: "send  SLIDES", due_date: null }, list);
  assert.strictEqual(view.mode, "source", "switched out of reading view");
  assert.deepStrictEqual(calls[0], ["activate", "note"]);
  assert.deepStrictEqual(calls[1], ["select", { line: 2, ch: 0 }, { line: 2, ch: text[2].length }]);
  calls.length = 0;
  await plugin.showOrInsert({ id: "2", title: "Something new", due_date: null }, list);
  assert.deepStrictEqual(text, ["# Day", "", "- [ ] Send slides 📅 2026-11-20", "    - [x] Book venue", "    - [ ] Something new"]);
  assert.ok(calls.some((c) => c[0] === "cursor" && c[1].line === 4), "cursor moves to the new line");
});

// A tiny editor over an array of lines, with the drop's selection.
function fakeEditor(lines, from, to) {
  let text = lines.join("\n");
  let selection = { anchor: from, head: to };
  const offset = (pos) => text.split("\n").slice(0, pos.line).reduce((n, l) => n + l.length + 1, 0) + pos.ch;
  return {
    get lines() {
      return text.split("\n");
    },
    get selection() {
      return selection;
    },
    listSelections: () => [selection],
    getLine: (n) => text.split("\n")[n],
    getRange: (a, b) => text.slice(offset(a), offset(b)),
    transaction({ changes }) {
      for (const c of changes) text = text.slice(0, offset(c.from)) + c.text + text.slice(offset(c.to));
    },
    setSelection: (anchor, head) => (selection = { anchor, head }),
  };
}

test("a task dropped inside a line moves to a line of its own", () => {
  const task = "- [ ] Send slides";
  // Dropped mid-sentence: the sentence stays whole, the task goes below it.
  let ed = fakeEditor([`Call Ann ${task}about the budget`, "next"], { line: 0, ch: 9 }, { line: 0, ch: 9 + task.length });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, ["Call Ann about the budget", task, "next"]);
  assert.deepStrictEqual(ed.selection, { anchor: { line: 1, ch: 0 }, head: { line: 1, ch: task.length } });
  // Dropped at the start of a line: the task goes above it.
  ed = fakeEditor([`${task}First line`], { line: 0, ch: 0 }, { line: 0, ch: task.length });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, [task, "First line"]);
  // Dropped at the end of a line.
  ed = fakeEditor([`Notes:${task}`], { line: 0, ch: 6 }, { line: 0, ch: 6 + task.length });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, ["Notes:", task]);
  // Dropped on an indented list item: the task gets the same indentation.
  ed = fakeEditor(["- Project", `    - Call Ann${task}`], { line: 1, ch: 14 }, { line: 1, ch: 14 + task.length });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, ["- Project", "    - Call Ann", `    ${task}`]);
  assert.deepStrictEqual(ed.selection, { anchor: { line: 2, ch: 4 }, head: { line: 2, ch: 4 + task.length } });
  // Dropped inside the indentation of a line: above it, same indentation.
  ed = fakeEditor([`\t${task}\t- Item`], { line: 0, ch: 1 }, { line: 0, ch: 1 + task.length });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, [`\t\t${task}`, "\t\t- Item"], "the line was indented by two tabs before the drop");
  // Dropped on an empty line: left alone.
  ed = fakeEditor(["a", task, "b"], { line: 1, ch: 0 }, { line: 1, ch: task.length });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, ["a", task, "b"]);
  // The selection is not the dropped text (something else happened): left alone.
  ed = fakeEditor(["Call Ann about it"], { line: 0, ch: 0 }, { line: 0, ch: 4 });
  placeOnOwnLine(ed, task);
  assert.deepStrictEqual(ed.lines, ["Call Ann about it"]);
});

test("the in-note filter shows only the tasks the active note holds", async () => {
  const plugin = await signedInPlugin();
  const view = global.__viewFactory({});
  view.lists = [
    { id: "L1", name: "Inbox", tasks: [{ id: "a", title: "Send slides", due_date: null }, { id: "b", title: "Book venue", due_date: null }] },
    { id: "L2", name: "Work", tasks: [{ id: "c", title: "Other", due_date: null }] },
  ];
  view.noteLines = noteTaskLines("- [ ] Send slides");
  const titles = () => view.groups().flatMap((g) => g.items.map((e) => e.task.title));
  assert.deepStrictEqual(titles(), ["Send slides", "Book venue", "Other"]);
  plugin.settings.onlyInNote = true;
  plugin.settings.tab = "lists";
  assert.deepStrictEqual(titles(), ["Send slides"]);
  assert.deepStrictEqual(view.groups().map((g) => g.name), ["Inbox"], "lists without a match are hidden");
  plugin.settings.tab = "dates";
  assert.deepStrictEqual(view.groups().map((g) => g.name), ["No date"], "empty date sections are hidden too");
  assert.strictEqual(view.inNoteIds(), "a");
});

test("links in task notes are found, without trailing punctuation", () => {
  const parts = splitLinks("See https://example.com/a?b=1. Also (https://x.org/y) and https://w.io/p_(1)!");
  assert.deepStrictEqual(
    parts.filter((p) => p.url).map((p) => p.url),
    ["https://example.com/a?b=1", "https://x.org/y", "https://w.io/p_(1)"]
  );
  assert.strictEqual(parts.map((p) => p.text).join(""), "See https://example.com/a?b=1. Also (https://x.org/y) and https://w.io/p_(1)!");
  assert.deepStrictEqual(splitLinks("no links"), [{ text: "no links" }]);
});

test("HTML task notes become plain text", async () => {
  const plugin = await signedInPlugin();
  global.__handler = (req) => {
    if (req.url.endsWith("/me/todo/lists")) return { body: { value: [{ id: "L1", displayName: "Inbox" }] } };
    return {
      body: {
        value: [
          { id: "t", title: "T", status: "notStarted", body: { contentType: "html", content: '<p>Call <b>Ann</b> &amp; Bo</p><p>Room&nbsp;4 &lt;3 <a href="https://example.com/doc">the doc</a></p>' } },
        ],
      },
    };
  };
  const lists = await plugin.fetchLists();
  assert.strictEqual(lists[0].tasks[0].notes, "Call Ann & Bo\nRoom 4 <3 the doc (https://example.com/doc)");
});

test("auth errors are trimmed and get a setup hint", () => {
  const message = describeAuthError({
    error_description: "AADSTS50059: No tenant-identifying information found. Trace ID: abc Correlation ID: def",
  });
  assert.ok(message.startsWith("AADSTS50059: No tenant-identifying information found."));
  assert.ok(!message.includes("Trace ID"));
  assert.ok(message.includes("Set Tenant to your directory (tenant) ID"));
});

test("note matching ignores case, indentation, whitespace and trailing emojis, dates and links", () => {
  const lines = noteTaskLines(
    [
      "- [x] Send slides 📅 2026-11-20",
      "    - [ ]   BOOK   the venue ✅ 16/11/2026 [To Do](https://example.com)",
      "- [>] Forwarded task",
      "- [ ] Release 2026 roadmap",
    ].join("\n")
  );
  const state = (title) => {
    const match = lines.titles.get(normalizeText(title));
    return match ? (match.checked ? "checked" : "open") : null;
  };
  assert.strictEqual(state("Send slides"), "checked");
  assert.strictEqual(state("book the venue"), "open");
  assert.strictEqual(state("Forwarded task"), "open", "[>] is not checked");
  assert.strictEqual(state("Release 2026 roadmap"), "open", "a year inside the title stays");
  assert.strictEqual(state("Send"), null, "no partial matches");
});

(async () => {
  let failed = 0;
  for (const [name, fn] of tests) {
    try {
      await fn();
      console.log(`ok   ${name}`);
    } catch (error) {
      failed++;
      console.log(`FAIL ${name}\n     ${error.stack.split("\n").slice(0, 3).join("\n     ")}`);
    }
  }
  console.log(`\n${tests.length - failed} passed, ${failed} failed`);
  process.exit(failed ? 1 : 0);
})();
