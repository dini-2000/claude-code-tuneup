# Setting up the handover scripts

The two scripts in `scripts/` put the rules in SKILL.md into practice. They are starting points: the user's AI reads this guide, fits them to the user's folders and backup repository, and checks with the user at each step. They hold no paths from any machine, since each finds its workspace from where it sits.

- `handover.py` lists the work across all projects: unfinished tasks, work held for later, notes from other agents, Status sections that need trimming, and the projects by their last handover. As a Claude Code SessionStart hook it gives that list to each new session, points a session that follows a `/clear` at its project's handover files, and asks a session that has been compacted to re-read IN PROGRESS.md. It only reads files.
- `backup.py` keeps the workspace in a private git repository: `start` before work, `push` when a task is finished, `save` when one is halted, `all` when usage is about to run out, and `check` to see what the next backup would carry. Before each commit it looks for secrets and for files too big for the host, and stops if it finds any. It never force-pushes.

Both need Python 3.8 or later and nothing outside its standard library; `backup.py` also needs git. On Windows the command is often `py -3` or `python` in place of `python3`.

## 1. Agree on the layout with the user
Ask, in one message:
- Which folder holds the projects. This is the workspace; each project is a folder inside it with a CLAUDE.md (or AGENTS.md).
- Where the backup should go: a new, private repository on any git host, such as GitHub, GitLab or Bitbucket, or a bare repository on another drive. It must be private, since it will hold everything in the workspace.
- Whether several accounts, machines or agents will work in the workspace, and what to call each one. The name signs each backup.

The user signs in to the host themselves, with an SSH key or the host's own credential helper (on GitHub, `gh auth login`). Never ask for a token or password, and never put one in a script, a settings file or the repository's URL, where git would keep it in plain text.

If the workspace is already a git repository, or sits inside one, ask the user how to proceed. Never change an existing repository's remote or history.

## 2. Copy the scripts and fit them
Copy `handover.py` and `backup.py` into a `_handover` folder in the workspace. From there each finds the workspace as the folder above its own, so no path needs changing. If the user wants them elsewhere, set `WORKSPACE` near the top of both to the workspace's full path.

Then read the settings block at the top of each and change only what the user's layout needs:
- `handover.py`: `DEPTH` (how many folder levels down projects can sit), `SKIP` (folder names that never hold projects), `HANDOVER_FILES` (put `AGENTS.md` first if another agent does most of the work), `SHOW_PROJECTS` and `ONLY_INSIDE`.
- `backup.py`: `REMOTE` and `BRANCH`; `MAX_MB`, the largest file the host takes (GitHub refuses files over 100 MB); `SECRETS` and `SECRET_NAMES`, if the user works with kinds of keys the lists lack.

Tell the user what you changed.

## 3. Make the workspace a git repository
On the first machine, in the workspace folder:

```bash
git init -b main
```

Add a `.gitignore` (start from the one below and ask the user what else should stay on this machine only), then connect the backup and look at what the first backup would carry:

```bash
git remote add origin <the private repository's URL>
python3 _handover/backup.py check
```

Go through the list with the user. Large media, other people's files that don't belong on a host, and anything else the user wants kept off the internet go in `.gitignore`. A project that is a git repository of its own, such as a code project with its own remote, goes in as a link only; its own remote is its backup. When the list looks right:

```bash
python3 _handover/backup.py all "first backup"
```

On a second machine, clone the backup to the workspace's path (`git clone <URL> <workspace>`) in place of `git init`, then add the hook and rules from steps 4 and 5. Accounts that share one machine share one workspace and need nothing more.

A `.gitignore` to start from:

```gitignore
# Secrets
.env
.env.*
!.env.example
*.p12
*.pfx
id_rsa*
id_ed25519*

# System and tool clutter
.DS_Store
Thumbs.db
__pycache__/
node_modules/
.venv/
venv/

# Lock and temporary files
~$*
.~lock.*#
*.tmp
```

## 4. Add the SessionStart hook
Make a copy of the user's Claude Code settings file (`~/.claude/settings.json`, which applies to every session), then add this entry, with the workspace's full path in place of `<workspace>`:

```json
{
  "hooks": {
    "SessionStart": [
      {
        "hooks": [
          {"type": "command", "command": "python3 \"<workspace>/_handover/handover.py\" hook"}
        ]
      }
    ]
  }
}
```

Merge it with what the file holds: keep every other setting and hook, and if `SessionStart` already has entries, add this one beside them. Forward slashes work in the path on Windows too. With no `matcher`, the hook runs at every session start: a new session, a resumed one, one after `/clear` and one after a compaction. Sessions that start outside the workspace get nothing from it.

An agent without hooks can run `handover.py status` at the start of each session; say so in its rules (step 5).

If each account or machine has a settings file of its own, add `"env": {"HANDOVER_NAME": "<name>"}` to each, and `backup.py` signs its commits with that name and the session id. Where accounts share one settings file, leave it out, and the session id signs alone.

## 5. Write the rules where the AI reads them
The scripts do nothing until the AI knows when to run them. Add a section like this to the workspace's CLAUDE.md (or AGENTS.md), which every session started inside the workspace reads, with the workspace's full path in place of `<workspace>`:

```markdown
## Handover and backup
Follow the agent-handover skill for what goes in each project's CLAUDE.md, IN PROGRESS.md, NOTES.md and HANDOFF.md. Below, `backup.py` stands for `python3 "<workspace>/_handover/backup.py"`.
- Before working on a project, run `backup.py start "<project folder>"`. If it stops, tell the user before working on files that may be out of date.
- When a task is finished and its handover is folded into CLAUDE.md: `backup.py push "<project folder>" "<what changed>"`.
- When the user halts a task, or usage or context nears a limit: bring IN PROGRESS.md up to date, then `backup.py save "<project folder>" "<where it stands>"`.
- When every account is near its limit, or before a long break: `backup.py all "<why>"`.
- If a backup stops on a secret or a large file, show the user the list and ask. Never force the file in, and never put a secret in the repository.
- Sign IN PROGRESS.md and handover log lines with your account name and the first 8 characters of the session id.
```

Add the `# Compact instructions` block from SKILL.md to the same file.

## 6. Test it with the user
- `python3 _handover/handover.py status` lists the projects. If it finds none, check `DEPTH` and that each project has a CLAUDE.md.
- Start a new session in the workspace and ask what the handover status says. Then run `/clear` in a session inside a project: the next session is told to read that project's CLAUDE.md.
- Run `backup.py start "<project>"`, change a file, run `backup.py push "<project>" "test"`, and find the commit on the host.
- To see the secret check work, create an empty file named `.env` in a project, run `backup.py check`, then delete the file.

## Ideas to build next
Add one only when the user wants it:
- Claims: `start` writes a claim file naming the session, and another session that finds a fresh claim stops and asks the user, so two sessions never edit one project at once.
- A daily backup: run `backup.py all "daily"` from the system's scheduler (cron, launchd or Task Scheduler), or from the hook when the last backup is more than a day old.
- A digest: after a cut-off, read the old session's transcript from the last handover save onwards to find what it did but never reported.
- Chat names: one name per project, the same on every account, so the user can say which chat to continue.
- A save before limits: the save-before-limit setting in Claude Code Tune-up asks the session to update IN PROGRESS.md before the chat compacts or a usage limit ends it.

## Undoing it
Remove the hook entry from the settings file and the section from CLAUDE.md, and delete `_handover/`. The workspace's git repository and the backup on the host stay until the user removes them.
