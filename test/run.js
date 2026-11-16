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
const mainPath = path.join(__dirname, "..", "src", "main.js");
const source =
  fs.readFileSync(mainPath, "utf8") +
  "\nmodule.exports.__internals = { MicrosoftAuth, noteTaskLines, normalizeText, describeAuthError };\n";
const mod = new Module(mainPath, module);
mod.filename = mainPath;
mod.paths = Module._nodeModulePaths(path.dirname(mainPath));
mod._compile(source, mainPath);
const TodoPanelPlugin = mod.exports;
const { MicrosoftAuth, noteTaskLines, normalizeText, describeAuthError } = TodoPanelPlugin.__internals;

global.window = { setTimeout: (fn) => setImmediate(fn), clearTimeout() {} };
global.atob = global.atob || ((s) => Buffer.from(s, "base64").toString("binary"));

const SESSION = "todo-panel-microsoft-session";
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
  assert.ok(plugin.app.secrets[SESSION].includes("RT1"));
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
  plugin.app.secrets[SESSION] = JSON.stringify({ client_id: "other", tenant: "tenant", refresh_token: "r" });
  assert.strictEqual(new MicrosoftAuth(plugin).signedIn, false);
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
  assert.ok(view.loading, "fetch running in the background");
  finishFetch([]);
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
  const state = (title) => (lines.titles.has(normalizeText(title)) ? (lines.titles.get(normalizeText(title)) ? "checked" : "open") : null);
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
