# Notes from the 2026-10-19 episode

- The Downloads zip was the pre-edit version. Repo files had been replaced by the zip contents and an extra Windows UTF-8 patch was reapplied on top of them.
- Git prints CRLF warnings on Windows. They are harmless.
- Remove `__pycache__` and local `.claude` folders before committing.
- Selftests that need LibreOffice fail on machines without it. Report the pass count together with the environment reason, not as a plain failure.
