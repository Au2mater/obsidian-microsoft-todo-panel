# To Do Panel for Obsidian

A sidebar in Obsidian that lists your Microsoft To Do tasks. Drag a task into
a note to insert it as a Markdown task line, or click it to insert it at the
cursor.

Microsoft To Do itself does not let you drag tasks into other apps, so the
drag starts in this panel instead.

- **Lists** and **By date** tabs: To Do's own lists, or Today, Later (earliest
  first), Earlier (earliest first) and No date. With a daily note open, the
  sections are relative to that note's date, and Today is titled with the
  date instead (for example "Thu 12 Nov") unless it is today's note. The date
  is read from the file name using the Daily notes plugin's format.
- One button expands or collapses every group.
- A task the active note already contains is marked **○ In note**, or
  **✓ In note** once a matching line is ticked with `[x]` (`[<]`, `[>]` and
  other marks do not count). Matching ignores case, indentation, whitespace,
  the checkbox, and trailing emojis, dates and links. Each group's title
  shows how many of its tasks the note holds, such as `○2 ✓1`, open or
  collapsed.
- A compact toggle shows recently completed tasks, struck through, after the
  open tasks of each group. Dropping one inserts `- [x]`.

The plugin only reads your tasks. It never creates, changes or completes
anything in To Do.

## Setup

The plugin signs in to Microsoft itself, with no other software needed. It
needs an app registration in Microsoft Entra so that Microsoft knows which
app is asking. There is no shared one: you register your own, once, which
takes about five minutes and keeps your tasks between you and Microsoft.

### 1. Register an app

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com)
   (or the Azure portal) and go to **App registrations → New registration**.
2. Name it, for example "Obsidian To Do Panel".
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

1. Open the panel and select **Sign in**, or open **Settings → To Do Panel**.
2. Paste the client ID into **Client ID**.
3. Set **Tenant**:
   - your directory (tenant) ID for an app limited to your organisation;
   - `consumers` for personal accounts only;
   - `common` for an app that accepts both work and personal accounts.
4. Select **Sign in**. A dialog shows a short code. Open the Microsoft page,
   enter the code and sign in. The dialog closes by itself and the panel fills.

Open the panel from the ribbon icon or with the command **To Do Panel: Open
panel**.

### Where the sign-in is kept

The sign-in (a refresh token) is stored in Obsidian's secret storage on this
device, never in the vault or the plugin's settings file, so it is not synced
or committed. Sign in once on each device. Changing the client ID or tenant,
or selecting **Sign out**, removes it.

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

Plain JavaScript, no build step. Copy `manifest.json`, `main.js` and
`styles.css` into `<vault>/<config folder>/plugins/second-brain-todo-panel/`,
then enable **To Do Panel** under **Settings → Community plugins**. The config
folder is `.obsidian` unless the vault overrides it.

```bash
./install.sh "/path/to/vault/.obsidian"
```

Requires Obsidian 1.11.4 or newer, for secret storage.

## Settings

| Setting | Default | Notes |
| --- | --- | --- |
| Client ID | empty | From your app registration. |
| Tenant | `common` | Tenant ID, `common`, `organizations` or `consumers`. |
| Permissions | `Tasks.Read` | Microsoft Graph permissions to ask for. |
| Completed tasks to load (days) | `7` | How far back the show-completed toggle reaches. |
| Task template | `- [ ] {{title}} {{due}}` | `{{title}}`, `{{due}}`, `{{date}}`, `{{list}}`, `{{link}}`, `{{id}}` |
| Due format | `📅 {{date}}` | `{{date}}` is YYYY-MM-DD. |
| Link format | `https://to-do.office.com/tasks/id/{{id}}/details` | `{{id}}`, `{{list}}` |

`{{due}}` is empty for a task with no due date, and the spaces around it
collapse.

## Development

```bash
node test/run.js
```

The tests load `main.js` against a small stand-in for the Obsidian API and
cover sign-in, token refresh, reading from Microsoft Graph and note matching.
