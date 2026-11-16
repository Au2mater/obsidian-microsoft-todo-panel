// A stand-in for the "obsidian" module, enough to load main.js in Node.
// requestUrl answers through global.__handler and records every request.
class Plugin {
  constructor() {
    const store = {};
    this.manifest = { id: "microsoft-todo-panel" };
    this.app = {
      secretStorage: {
        setSecret(id, value) {
          if (!/^[a-z0-9-]+$/.test(id)) throw new Error("Invalid secret ID");
          store[id] = value;
        },
        getSecret: (id) => store[id] ?? null,
        listSecrets: () => Object.keys(store),
      },
      workspace: {
        on: () => ({}),
        getLeavesOfType: () => [],
        getActiveFile: () => null,
        getActiveViewOfType: () => null,
        getMostRecentLeaf: () => null,
        setActiveLeaf() {},
        onLayoutReady: (fn) => (global.__layoutReady = fn),
      },
      vault: { on: () => ({}), getName: () => "Test vault" },
      secrets: store,
      local: {},
      loadLocalStorage(key) {
        return this.local[key] ?? null;
      },
      saveLocalStorage(key, value) {
        if (value === null) delete this.local[key];
        else this.local[key] = JSON.parse(JSON.stringify(value));
      },
    };
  }
  registerEvent() {}
  registerInterval(id) {
    clearInterval(id);
    return id;
  }
  registerView(type, factory) {
    global.__viewFactory = factory;
  }
  addRibbonIcon() {}
  addCommand() {}
  addSettingTab() {}
  async loadData() {
    return global.__data || {};
  }
  async saveData(data) {
    global.__saved = data;
  }
}

class Stub {}

module.exports = {
  ItemView: class {
    constructor(leaf) {
      this.leaf = leaf;
    }
  },
  MarkdownView: Stub,
  Modal: Stub,
  Notice: class {
    constructor(message) {
      (global.__notices = global.__notices || []).push(message);
    }
  },
  Plugin,
  PluginSettingTab: Stub,
  Setting: Stub,
  setIcon() {},
  moment: () => ({ isValid: () => false, format: () => "" }),
  async requestUrl(req) {
    global.__requests.push(req);
    const res = global.__handler(req);
    return {
      status: res.status || 200,
      headers: res.headers || {},
      get json() {
        if (res.body === undefined) throw new Error("Not JSON");
        return res.body;
      },
    };
  },
};
