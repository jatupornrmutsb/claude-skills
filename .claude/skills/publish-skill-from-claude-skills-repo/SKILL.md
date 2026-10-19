---
name: publish-skill-from-claude-skills-repo
description: Use when publishing a skill edit to claude-skills repo.
category: git-workflow
---

# Publish skill edits from the claude-skills repo

The repo is the source of truth. Edit skills in the repo clone, never only in `~/.claude/skills/` or in an uploaded zip. A Desktop-only edit is lost on the next repo sync.

After every skill edit:
1. Commit and push to `origin master`.
2. Wait until the `skill-zips.yml` workflow succeeds for that exact commit. GitHub then rebuilds the zips on the `skills-latest` release.
3. Copy each changed skill folder into `~/.claude/skills/` (replace the folder, do not merge).
4. If a file under `commands/` changed, copy it to `~/.claude/commands/` too. Skill copies do not update commands.
5. Tell the user what to do in Claude Desktop: delete the old skill under Settings → Capabilities → Skills, upload the new zip from Releases, and confirm its toggle is on.

Do not tell the user the zip is updated until step 2 has succeeded. When the user uploads a zip from Downloads, diff it against the repo first and say what differs before overwriting.

Details: `references/workflow-notes.md`.
