---
name: block-gratuitous-assert
enabled: true
event: file
file_pattern: \\.py$
pattern: assert\\s+\\w+\\s+is\\s+not\\s+None
action: warn
---

**Gratuitous assert detected**

`assert x is not None` just to satisfy the type checker is usually a code smell.

Better approaches:
- Add an early return: `if x is None: return`
- Use a conditional: `if x is not None: x.method()`
- Fix the upstream function to not return Optional when it shouldn't
- Raise a proper exception with context if None is truly unexpected

Only use assert when None genuinely indicates a programming error, not to suppress type warnings.
