---
name: context-audit
description: Audits what Claude Code loads into every session (CLAUDE.md files, the auto-memory index, connectors and their tools) and trims it without losing anything. Run it when chats compact early, feel slow or burn through usage limits.
disable-model-invocation: true
---
# Context audit

CLAUDE.md files, the auto-memory index and the tool list of every connected server are in every session from its first request, and every later request sends them again. Trimming them pays off in every chat.

## 1. Measure
- `/context` shows the biggest categories in the current window; `/usage` shows where plan usage goes.
- `wc -c` each CLAUDE.md in scope (user, project, subfolders) and the memory index, `MEMORY.md` in the project's memory folder. `awk 'length > 200' MEMORY.md | wc -l` counts long index lines.
- `/mcp` lists the connected servers and their tools.

## 2. CLAUDE.md: live content stays, history moves
- Keep what a new session needs without opening another file: current status, live decisions in the user's words, file locations, how to build and check, conventions, known pitfalls, the last few log lines.
- Move history, word for word, to a `NOTES.md` beside it: dated status entries, older log lines, superseded decisions, long reference tables. Leave one line saying what moved and when to read it.
- Before moving a dated entry, check that its live facts (paths, commands, settings) also appear in a section that stays. If not, keep the entry.
- Never reword or trim live content to save space.

## 3. Memory index: one short line each
- Each index line is a title, a link and a short hook, under about 200 characters. The detail lives in the topic file.
- Before shortening a line, append the old one word for word to its topic file. Keep any constraint the line carries, such as "never mention X".

## 4. Connectors: off where unused
- Every connected server's tools ride along with each request. Ask the user which connectors they never use from Claude Code and suggest turning those off for Claude Code only. Change connector settings only with the user's yes.

## 5. Report
Sizes before and after, what moved where, and anything left for the user.
