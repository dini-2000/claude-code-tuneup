#!/usr/bin/env python3
"""Handover status for a folder of projects, and a Claude Code SessionStart hook that shows it.

    python3 handover.py status      print the status
    python3 handover.py hook        run as a SessionStart hook (reads the hook's JSON on stdin)

A project is any folder below the workspace that holds a CLAUDE.md, an AGENTS.md or an IN PROGRESS.md (SKILL.md says
what each file holds). The status lists unfinished tasks, work held until the user says go, notes from other agents
waiting to be folded in (HANDOFF.md), Status sections long enough to need trimming, and the projects by their last
handover.

As a hook it gives that status to every new or resumed session that starts inside the workspace. After /clear it
points the new session at its project's handover files and at the transcript of the session before the clear; after a
compaction it asks for IN PROGRESS.md to be read again. It reads files only and never writes.

This is a starting point: setup.md says how to fit it to a user's folders. Python 3.8 or later, standard library only.
"""
import datetime
import json
import re
import sys
from pathlib import Path

# Settings: adapt these during setup (setup.md, step 2).
WORKSPACE = Path(__file__).resolve().parent.parent  # this script sits in <workspace>/_handover/
DEPTH = 3  # folder levels below the workspace searched for projects
SKIP = {'node_modules', '__pycache__', 'venv', '_archive', '_handover', 'old', 'build', 'source', 'correspondence'}
HANDOVER_FILES = ('CLAUDE.md', 'AGENTS.md')  # the first one found is the project's handover
PROGRESS = 'IN PROGRESS.md'
STATUS_FLAG = 8000  # flag a Status section longer than this many characters
SHOW_PROJECTS = 10  # projects listed by last handover in the hook (the status command lists all)
CONTEXT_CAP = 9500  # Claude Code keeps at most 10,000 characters of a hook's additionalContext
ONLY_INSIDE = True  # the hook stays quiet in sessions that start outside the workspace


def read(path):
    try:
        return path.read_text(encoding='utf-8', errors='replace')
    except OSError:
        return ''


def stamp(path, fmt='%d %b %H:%M'):
    return datetime.datetime.fromtimestamp(path.stat().st_mtime).strftime(fmt)


def is_project(folder):
    return any((folder / name).is_file() for name in HANDOVER_FILES + (PROGRESS,))


def find_projects():
    """Every project folder below the workspace, down to DEPTH levels; hidden folders and SKIP are left out."""
    found = []

    def walk(folder, depth):
        if depth > DEPTH:
            return
        try:
            subs = sorted(p for p in folder.iterdir() if p.is_dir())
        except OSError:
            return
        for sub in subs:
            if sub.name.startswith('.') or sub.name in SKIP:
                continue
            if is_project(sub):
                found.append(sub)
            walk(sub, depth + 1)

    walk(WORKSPACE, 1)
    return found


def handover_file(folder):
    return next((folder / n for n in HANDOVER_FILES if (folder / n).is_file()), None)


def status_section(text):
    m = re.search(r'^## Status[^\n]*\n(.*?)(?=^## |\Z)', text, re.S | re.M)
    return m.group(1) if m else ''


def now_line(text):
    m = re.search(r'^Now:\s*(.+)$', status_section(text) or text, re.M)
    return m.group(1).strip() if m else ''


def progress(folder):
    """Title and Held: line of the project's IN PROGRESS.md, or None when no task is under way."""
    path = folder / PROGRESS
    if not path.is_file():
        return None
    lines = [line.strip() for line in read(path).splitlines() if line.strip()]
    title = lines[0].lstrip('#').strip() if lines else '(empty file)'
    held = next((line for line in lines[1:6] if line.startswith('Held:')), '')
    return {'title': title, 'held': held, 'updated': stamp(path), 'path': path}


def name_of(folder):
    return folder.relative_to(WORKSPACE).as_posix()


def short(text, n=160):
    return text if len(text) <= n else text[:n - 3].rstrip() + '...'


def status_text(limit=None):
    tasks, held, notes, long_status, handovers = [], [], [], [], []
    for folder in find_projects():
        name = name_of(folder)
        task = progress(folder)
        if task and task['held']:
            held.append('- %s: %s\n  %s' % (name, short(task['title']), short(task['held'])))
        elif task:
            tasks.append('- %s: %s (updated %s)' % (name, short(task['title']), task['updated']))
        if any(folder.glob('HANDOFF*.md')):
            notes.append('- ' + name)
        path = handover_file(folder)
        if path:
            text = read(path)
            size = len(status_section(text))
            if size > STATUS_FLAG:
                long_status.append('- %s (%dK)' % (name, size // 1000))
            handovers.append((path.stat().st_mtime, name, stamp(path, '%d %b'), now_line(text)))

    when = datetime.datetime.now().strftime('%d %b %Y %H:%M')
    if not (tasks or held or handovers):
        return ('Handover status, %s: no projects found below %s. A project is a folder with a %s or an %s.'
                % (when, WORKSPACE, ' or '.join(HANDOVER_FILES), PROGRESS))
    out = ['Handover status, %s (workspace %s).' % (when, WORKSPACE)]
    if tasks:
        out += ['Unfinished tasks (%s): pick up only this session\'s own project.' % PROGRESS] + tasks
    if held:
        out += ['Held until the user says go (neither start nor route them):'] + held
    if notes:
        out += ['Notes from other agents to read with the handover and fold in (HANDOFF.md):'] + notes
    if long_status:
        out += ['Status sections past {:,} characters (move dated entries to NOTES.md word for word at the next '
                'handover):'.format(STATUS_FLAG)] + long_status
    handovers.sort(reverse=True)
    shown = handovers if limit is None else handovers[:limit]
    if shown:
        out.append('Projects by last handover%s:' % ('' if len(shown) == len(handovers)
                                                     else ', newest %d of %d' % (len(shown), len(handovers))))
        out += ['- %s (%s)%s' % (name, day, ': ' + short(now) if now else '') for _, name, day, now in shown]
    return '\n'.join(out)


def resolved(cwd):
    try:
        return Path(cwd).resolve()
    except (OSError, RuntimeError, ValueError):
        return None


def inside_workspace(cwd):
    here = resolved(cwd)
    return here is not None and (here == WORKSPACE or WORKSPACE in here.parents)


def nearest_project(cwd):
    """The project folder at or above cwd, inside the workspace."""
    here = resolved(cwd)
    if here is None:
        return None
    for folder in (here, *here.parents):
        if folder == WORKSPACE or WORKSPACE not in folder.parents:
            return None
        if is_project(folder):
            return folder
    return None


def previous_transcript(current):
    """Claude Code keeps each session's transcript in one folder per working directory, so the newest other
    transcript there is most likely the session before the clear."""
    if not current:
        return None
    current = Path(current)
    try:
        others = [p for p in current.parent.glob('*.jsonl') if p.name != current.name]
        return max(others, key=lambda p: p.stat().st_mtime, default=None)
    except OSError:
        return None


def hook_text(event):
    source = event.get('source') or 'startup'
    if ONLY_INSIDE and not inside_workspace(event.get('cwd') or '.'):
        return ''
    project = nearest_project(event.get('cwd') or '.')
    if source == 'clear' and project:
        handover = handover_file(project)
        task = progress(project)
        text = ('Handover: this session follows a /clear. Before the next task, read %s to the end (a long file can '
                'take more than one read)%s.' % (handover or 'the project\'s handover file',
                                                   ', then %s' % task['path'] if task else ''))
        before = previous_transcript(event.get('transcript_path'))
        if before:
            text += (' The session before the clear is most likely %s: if the user refers to something the handover '
                     'lacks, read what that session did after the handover was last saved.' % before)
        return text
    if source == 'compact':
        task = progress(project) if project else None
        text = 'Handover: this chat was just compacted, so its older turns are now a summary. Before carrying on, '
        if task:
            return text + ('re-read %s and the part of the project handover the next step needs. If it is behind the '
                           'summary\'s current step, update it from the summary.' % task['path'])
        return text + ('re-read the part of the project handover the next step needs. If a task is under way and has '
                       'no %s yet, start one from the summary.' % PROGRESS)
    return status_text(limit=SHOW_PROJECTS)  # startup, resume, fork, and a clear outside any project


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')  # Windows consoles default to a narrower code page
    except (AttributeError, ValueError):
        pass
    command = sys.argv[1] if len(sys.argv) > 1 else 'status'
    if command == 'status':
        print(status_text())
    elif command == 'hook':
        try:
            event = json.load(sys.stdin)
        except ValueError:
            event = {}
        text = hook_text(event if isinstance(event, dict) else {})
        if not text:
            return
        if len(text) > CONTEXT_CAP:
            text = text[:CONTEXT_CAP].rsplit('\n', 1)[0] + '\n(cut to fit; run handover.py status for the rest)'
        print(json.dumps({'hookSpecificOutput': {'hookEventName': 'SessionStart', 'additionalContext': text}}))
    else:
        sys.exit(__doc__)


if __name__ == '__main__':
    main()
