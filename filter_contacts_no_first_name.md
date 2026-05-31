# Filter iPhone Contacts With No First Name (Apple Shortcuts)

## Problem

Using a **Filter Contacts** action with `First Name is (empty)` or similar returns zero results, even when contacts without a First Name exist.

**Root cause:** Contacts missing a First Name have a `null`/absent field — not an empty string. The Filter Contacts action cannot reliably match null field values in many iOS versions.

## Working Solution: Repeat Loop

Build this in the Shortcuts app:

```
1. Get All Contacts
2. Set Variable "No First Name" → (empty, no value)
3. Repeat with each item in Contacts
4.   Get Details of Contacts → First Name → from Repeat Item
5.   If [First Name result] does not have a value
6.     Add Repeat Item to Variable "No First Name"
7.   End If
8. End Repeat
9. Get Variable "No First Name"
```

The result of step 9 is your list of contacts with no First Name.

## Why Each Approach Fails or Works

| Approach | Result |
|---|---|
| Filter Contacts → First Name is "" | Returns nothing — matches empty string, not null |
| Filter Contacts → First Name does not have a value | Returns nothing — Filter block bug with null contact fields |
| Repeat loop + "does not have a value" on the extracted field | **Works** — evaluates null correctly at the individual field level |

## Notes

- "Get Details of Contacts → First Name" extracts just the first name field from each contact.
- The `does not have a value` check in an **If** block correctly handles null, unlike the Filter block.
- This approach is slower on large contact lists (it loops one by one) but is the only reliable method.
