# To Do Panel for Obsidian

A sidebar in Obsidian that lists your open Microsoft To Do tasks. Drag a task
into a note to insert it as a Markdown line, or click it to insert it at the
cursor.

Microsoft To Do itself does not allow dragging tasks into other apps, so the
drag starts in this panel instead.

## How it works

The plugin holds no sign-in of its own. It runs `second-brain todo-lists --json`,
which reads the lists and open tasks live from Microsoft 365 using the CLI's
cached sign-in. That command never prompts: when the sign-in has expired it
fails, and the panel shows the message. Run `second-brain sync` in a terminal
to sign in again, then refresh the panel.

By default the command runs inside WSL:

```text
wsl.exe -e bash -lc "second-brain todo-lists --json"
```

Change it under **Settings → To Do Panel → Command** if the CLI is installed
on Windows directly (for example `second-brain todo-lists --json`).

## Install

Plain JavaScript, no build step. Copy `manifest.json`, `main.js` and
`styles.css` into `<vault>/<config folder>/plugins/second-brain-todo-panel/`,
then enable **To Do Panel** under **Settings → Community plugins**. The config
folder is `.obsidian` unless the vault overrides it.

```bash
./install.sh "/path/to/vault/.obsidian"
```

It needs a second-brain CLI new enough to have the `todo-lists` command.

## Settings

| Setting | Default | Placeholders |
| --- | --- | --- |
| Task template | `- [ ] {{title}} {{due}}` | `{{title}}`, `{{due}}`, `{{date}}`, `{{list}}`, `{{link}}`, `{{id}}` |
| Due format | `📅 {{date}}` | `{{date}}` (YYYY-MM-DD) |
| Link format | `https://to-do.office.com/tasks/id/{{id}}/details` | `{{id}}`, `{{list}}` |

`{{due}}` is empty for a task with no due date, and the spaces around it
collapse.
