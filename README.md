# Microsoft To Do Panel for Obsidian

A sidebar in Obsidian that lists your Microsoft To Do tasks. Drag a task into
a note to insert it as a Markdown task line: it lands on a line of its own,
not inside an existing line, indented like the line you dropped it on. Or use
a task's **+** button to add it at the cursor: on the cursor's line if that is
empty, otherwise on a new line below it, with the same indentation, and never
inside the note's properties. Clicking a task the note already holds selects
its line there. Select **Notes** under a task to read its notes; links in them
open in your browser.

Microsoft To Do itself does not let you drag tasks into other apps, so the
drag starts in this panel instead.

- **Lists** and **By date** tabs: To Do's own lists, or Today, Later (earliest
  first), Earlier (earliest first) and No date. With a daily note open, the
  sections are relative to that note's date, and Today is titled with the
  date instead (for example "Thu 12 Nov") unless it is today's note. The date
  is read from the file name with the **Daily note date format** setting.
- One button expands or collapses every group.
- A task the active note already contains is marked **○ In note**, or
  **✓ In note** once a matching line is ticked with `[x]` (`[<]`, `[>]` and
  other marks do not count). Matching ignores case, indentation, whitespace,
  the checkbox, and trailing emojis, dates and links. Each group's title
  shows how many of its tasks the note holds, such as `○2 ✓1`, open or
  collapsed.
- A compact toggle shows only the tasks the active note holds, and follows
  you as you switch notes.
- A compact toggle shows recently completed tasks, struck through, after the
  open tasks of each group. Dropping one inserts `- [x]`.

The plugin only reads your tasks. It never creates, changes or completes
anything in To Do.

This plugin is not made by or affiliated with Microsoft. Microsoft To Do is a
trademark of Microsoft.

## Disclosures

- **Account required:** a Microsoft account (work, school or personal) with
  Microsoft To Do, and an app registration you create yourself (see Setup).
- **Network use:** the plugin talks only to Microsoft:
  `login.microsoftonline.com` to sign in and `graph.microsoft.com` to read
  your To Do lists and tasks. Nothing is sent anywhere else.
- **Clipboard:** the sign-in dialog writes the sign-in code to the clipboard
  when you select **Copy code**. The plugin never reads the clipboard.
- **No telemetry, ads or payments.** Task data is not written to your vault
  unless you drag a task into a note or use its + button.
- **Cached on your device:** the last fetched lists are kept in Obsidian's
  local storage for this vault (not in the vault's files), so the panel shows
  them instantly at startup while it refreshes. Signing out clears them.
- **Stored on your device:** the sign-in token, in Obsidian's secret storage
  (see below). The client ID, tenant and other settings are in the plugin's
  `data.json` in your vault.

## Setup

The plugin signs in to Microsoft itself, with no other software needed. It
needs an app registration in Microsoft Entra so that Microsoft knows which
app is asking. There is no shared one: you register your own, once, which
takes about five minutes and keeps your tasks between you and Microsoft.

### 1. Register an app

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com)
   (or the Azure portal) and go to **App registrations → New registration**.
2. Name it, for example "Obsidian Microsoft To Do Panel".
3. Under **Supported account types**, pick one:
   - **Accounts in this organizational directory only** for a work or school
     account. You will then use your tenant ID in the plugin.
   - **Personal Microsoft accounts only**, or **Any organizational directory
     and personal Microsoft accounts**, for an Outlook.com or Hotmail account.
4. Leave **Redirect URI** empty and select **Register**.
5. On the app's **Overview** page, copy the **Application (client) ID**, and
   for a work account the **Directory (tenant) ID**.
6. Under **Authentication**, set **Allow public client flows** to **Yes** and
   save. The plugin signs in with a device code, which needs this.
7. Under **API permissions**, select **Add a permission → Microsoft Graph →
   Delegated permissions**, add **Tasks.Read**, and save. If your organisation
   requires it, ask an administrator to grant consent.

### 2. Connect the plugin

1. Open the panel and select **Sign in**, or open **Settings → Microsoft To Do Panel**.
2. Paste the client ID into **Client ID**.
3. Set **Tenant**:
   - your directory (tenant) ID for an app limited to your organisation;
   - `consumers` for personal accounts only;
   - `common` for an app that accepts both work and personal accounts.
4. Select **Sign in**. The dialog shows the client ID and tenant to check;
   select **Continue** and it shows a short code. Open the Microsoft page,
   enter the code and sign in. The dialog closes by itself and the panel fills.
   Clicking outside the dialog does not close it, so you can switch to the
   browser and back; press Escape to close it.

Open the panel from the ribbon icon or with the command **Microsoft To Do Panel: Open panel**.

### Where the sign-in is kept

The sign-in (a refresh token) is stored in Obsidian's secret storage on this
device, never in the vault or the plugin's settings file, so it is not synced
or committed. Each vault keeps its own sign-in; sign in once per vault on each
device. Changing the client ID or tenant, or selecting **Sign out**, ends it.

### If sign-in fails

The dialog shows Microsoft's message, with a hint for the common setup
mistakes:

| Message starts with | Fix |
| --- | --- |
| `AADSTS50059` | Set **Tenant** to your directory (tenant) ID. |
| `AADSTS700016` | Check the client ID and tenant. |
| `AADSTS7000218` | Turn on **Allow public client flows** in the app registration. |
| `AADSTS65001` | The permission is not granted yet; an administrator may need to approve it. |

Some organisations block device code sign-in altogether. If yours does, the
plugin cannot sign in there.

## Install

Requires Obsidian 1.11.4 or newer, for secret storage.

- **From Obsidian:** search for "Microsoft To Do Panel" under **Settings →
  Community plugins → Browse**, install and enable it.
- **With BRAT:** add `Au2mater/obsidian-microsoft-todo-panel` in the BRAT
  plugin to get the latest release.
- **Manually:** download `main.js`, `manifest.json` and `styles.css` from the
  [latest release](https://github.com/Au2mater/obsidian-microsoft-todo-panel/releases/latest)
  into `<vault>/.obsidian/plugins/microsoft-todo-panel/`, then enable the
  plugin under **Settings → Community plugins**.

## Settings

| Setting | Default | Notes |
| --- | --- | --- |
The settings tab shows your Microsoft account and the task template. The rest
is under **Advanced**, collapsed.

| Setting | Default | Notes |
| --- | --- | --- |
| Client ID | empty | From your app registration. Shown up front until you sign in. |
| Tenant | `common` | Tenant ID, `common`, `organizations` or `consumers`. Shown up front until you sign in. |
| Task template | `- [ ] {{title}} {{due}}` | `{{title}}`, `{{due}}`, `{{date}}`, `{{list}}`, `{{link}}`, `{{id}}` |
| Due date format (advanced) | `📅 {{date}}` | What `{{due}}` becomes. `{{date}}` is YYYY-MM-DD. |
| Completed tasks to load (advanced) | `7` days | How far back the show-completed toggle reaches; 0 hides it. |
| Daily note date format (advanced) | `YYYY-MM-DD` | Moment.js format of daily note file names. |
| Permissions (advanced) | `Tasks.Read` | Microsoft Graph permissions to ask for. |

`{{due}}` is empty for a task with no due date, and the spaces around it
collapse. `{{link}}` is `https://to-do.office.com/tasks/id/<task id>/details`.
Lists with no tasks to show are hidden.

## Development

Plain JavaScript with no dependencies. The source is `src/plugin.js`;
`npm run build` checks it and writes `main.js`, the file Obsidian loads. To
try a change, build it into a vault:

```bash
./install.sh "/path/to/vault/.obsidian"   # builds, then copies
npm test
```

The tests load `src/plugin.js` against a small stand-in for the Obsidian API and
cover sign-in, token refresh, reading from Microsoft Graph and note matching.

To release, bump `version` in `manifest.json` and `package.json`, add it to
`versions.json` and commit. Push the tag first and the branch only once the
release exists, because the community directory reads `manifest.json` from the
default branch and expects a release for that version:

```bash
git tag 1.2.0 && git push origin 1.2.0   # the workflow builds and publishes
# wait for the release to appear, then:
git push origin main
```

The release workflow tests the plugin, attests the build provenance of the
release files and publishes the GitHub release.
