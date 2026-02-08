---
name: warn-any-type
enabled: true
event: file
file_pattern: \\.py$
pattern: ":\\s*Any\\b|->\\s*Any\\b|\\[Any\\]"
action: warn
---

**Explicit Any type detected**

Using `Any` defeats the purpose of type checking.

Prefer:
- Specific types when known
- TypeVar for generic functions
- Union types for multiple possibilities
- object for "any object" (still type-safe)
- Protocol for structural typing

Only use Any when interfacing with truly dynamic code (e.g., some third-party APIs).
