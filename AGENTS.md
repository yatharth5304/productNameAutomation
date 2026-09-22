\# Product Name Automation — Agent Instructions



\## Core Rule



Prioritize correctness and evidence over speed, confidence, or unnecessary exploration.



\## Tool-Use Efficiency



\- Before every tool call, check whether the required information is already available from a previous tool result in the current session.

\- NEVER repeat a command when its previous output is already available and the repository state has not changed.

\- NEVER repeat repository-discovery commands unnecessarily.

\- NEVER repeatedly run `Get-ChildItem`, directory listings, `git status`, `git log`, `git branch`, `grep`, `glob`, or equivalent commands when their relevant results are already known.

\- Do not run the same inspection through a different command merely to reconfirm information already established.

\- Group related checks into one command when practical.

\- Reuse previously obtained file contents and command output.

\- Avoid unnecessary inspect → reason → inspect cycles.



\## Repository Exploration



\- Inspect the repository structure once when discovery is required.

\- After identifying relevant files, stop broad repository exploration.

\- Read only files directly relevant to the current task.

\- Do not inspect unrelated files, dependencies, logs, backups, binaries, generated artifacts, or large datasets unless directly necessary.

\- Do not repeatedly reopen a file that has already been read successfully.

\- Do not inspect every file simply because it exists.



\## Expensive Operations



\- Do not run training, fine-tuning, model inference, benchmarks, builds, tests, or other expensive/long-running operations unless the task explicitly requires them.

\- Prefer existing recorded results over recomputing results.

\- Prefer lightweight metadata checks over loading large datasets unnecessarily.

\- Do not execute commands merely to confirm information that is already established.



\## Task Boundaries



\- Follow the user's explicit task scope exactly.

\- Do not perform additional work that was not requested.

\- If the user says inspection/audit only, do not modify files.

\- If the user says not to run a particular operation, do not run it.

\- Clearly distinguish verified facts from assumptions or inferences.

\- Never present an inferred value as a verified fact.

\- When information cannot be verified from the available repository evidence, say so explicitly.



\## Stopping Condition



\- Stop investigating once enough evidence has been gathered to complete the requested task accurately.

\- Do not continue exploring solely to increase confidence when the existing evidence is sufficient.

\- Do not perform additional repository discovery after the requested answer can already be supported by the evidence collected.



\## File Modification



\- Never modify `AGENTS.md` unless the user explicitly requests it.

\- Before modifying project files, ensure the current task actually requires modification.

\- Do not create unnecessary temporary, report, cache, or generated files.



\## Git Safety



\- Do not commit, reset, rebase, checkout, clean, delete, or otherwise alter Git history or working-tree state unless explicitly requested.

\- Do not repeatedly run Git discovery commands when the relevant result is already known.



\## Context Efficiency



\- Avoid unnecessarily loading large files or datasets into context.

\- When only a portion of a file is relevant, inspect only the relevant portion when practical.

\- Reuse context already obtained instead of reconstructing it.

\- Keep intermediate reasoning and tool usage focused on the requested objective.



\## General Behavior



\- Prefer the smallest number of tool calls that can reliably complete the task.

\- Optimize for fewer model turns and tool calls without sacrificing correctness.

\- Do not sacrifice required investigation merely to reduce request count.

