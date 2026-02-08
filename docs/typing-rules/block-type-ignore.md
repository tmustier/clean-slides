---
name: block-type-ignore
enabled: true
event: file
file_pattern: \\.py$
pattern: "#\\s*type:\\s*ignore"
action: block
---

**Type ignore comment blocked**

Do NOT use `# type: ignore` to suppress errors. Instead:
- Fix the actual type issue
- Add proper type annotations
- Refactor the code if the type system reveals a design problem

If genuinely needed (rare), document WHY in `.long-task-harness/typing-findings.md` first.
