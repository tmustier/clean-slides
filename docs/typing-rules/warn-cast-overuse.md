---
name: warn-cast-overuse
enabled: true
event: file
file_pattern: \\.py$
pattern: from typing import.*cast|cast\\(
action: warn
---

**Cast usage detected**

`cast()` tells the type checker "trust me" without runtime verification.

Prefer:
- Proper type narrowing with isinstance()
- TypeGuard functions for complex narrowing
- Fixing the upstream type annotations

Only use cast when you have verified the type is correct and cannot express it otherwise.
