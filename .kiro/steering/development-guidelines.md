# Development Guidelines

## Outlook / COM Automation

- Do NOT run or execute anything that invokes Outlook (or any Office COM automation) on the development machine.
- Treat Outlook and the Exchange/MAPI layer as a black box. We write code against its interface but never call it locally.
- Testing and validation should use mocks, stubs, or sample data rather than live COM connections.
- When verifying scripts, limit checks to syntax/compilation — do not attempt runtime execution of Outlook-dependent code.

## Git Workflow

- Create a git commit after every set of changes resulting from a single user prompt.
- Use a concise, descriptive commit message summarizing what was done.
- Stage only the files that were modified as part of that prompt's work.
