# Filter iPhone Contacts With No First Name (Apple Shortcuts)

## Root Cause

`Get First Name from Contacts` returns an **empty string `""`** when a contact has no first name — not `null`. So checking `does not have any value` always evaluates to `false`, because `""` is technically a value.

| Contact state | What the action returns | `does not have any value` |
|---|---|---|
| Has first name "John" | `"John"` | false |
| Has no first name | `""` (empty string) | **false** ← bug |

## Correct Solution

```
1.  Find Contacts (no filter)
2.  Repeat with each item in Contacts
3.    Get First Name from Repeat Item
4.    Count Characters in First Name
5.    If [Count] is 0
6.      Add Repeat Item to MissingName
7.    End If
8.  End Repeat
9.  Count Items in MissingName   ← use "Items", not "Characters"
10. Show alert [Count]
11. Show MissingName
```

The character count of `""` is `0`, so `If [Count] is 0` correctly catches both null and empty-string first names.

## What Was Wrong in the Original Shortcut

1. **`If First Name does not have any value` never triggered** — `Get First Name from Contacts` returns `""` (empty string) for missing first names, and `""` is not null, so the condition was always false.

2. **Redundant first loop** — The `Find Contacts → Repeat → Add to AllContacts` loop just copies all contacts into a variable for no reason. Iterate directly over the `Find Contacts` output instead.

3. **`Count Characters in MissingName`** — Counts characters across all contact name strings, not the number of contacts. Use `Count Items` to get the contact count.
