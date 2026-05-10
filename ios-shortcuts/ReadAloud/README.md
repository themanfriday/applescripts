# Read Aloud — iOS Shortcut

An iOS Shortcut that reads aloud text from any source:

| Source | How to trigger |
|--------|---------------|
| **Webpage** | Share a URL from Safari (or any browser) using the Share Sheet |
| **File** | Share a file (PDF, .txt, Pages doc, Word doc, etc.) using the Share Sheet |
| **Plain/rich text** | Share selected text from any app using the Share Sheet |
| **Clipboard** | Run the shortcut directly (from the Shortcuts app, Home Screen, or Siri) with nothing shared — it reads whatever is on the clipboard |

---

## Installation

### Option A — import the pre-built shortcut file

1. Copy `ReadAloud.shortcut` to your iPhone or iPad (AirDrop, iCloud Drive, or cable).
2. Tap the file in the Files app. iOS will ask whether to add it to Shortcuts — tap **Add Shortcut**.

### Option B — build it manually in the Shortcuts app

Follow the step-by-step instructions in the [Manual build](#manual-build) section below.

---

## How it works

```
Shortcut Input received?
│
├── YES: does it contain a URL?
│   ├── YES → fetch webpage → extract article body → speak
│   └── NO  → extract text from input (file or plain text) → speak
│
└── NO (run standalone) → read clipboard → speak
```

The webpage path uses Shortcuts' built-in article extractor (equivalent to
Safari Reader), so it reads the meaningful body text rather than navigation
menus and ads.

---

## Adjustable settings

Open the shortcut in the Shortcuts editor and find any **Speak Text** action to
change:

| Setting | Default | Notes |
|---------|---------|-------|
| Rate | 0.45 | 0 = slowest, 1 = fastest |
| Pitch | 1.0 | 0 = lowest, 2 = highest |
| Language | device default | e.g. `en-US`, `fr-FR` |
| Wait until done | on | turn off to run other shortcuts while speaking |

---

## Manual build

Use these steps to recreate the shortcut from scratch in the Shortcuts app
(iOS 16 or later).

### Shortcut settings (tap the ⓘ icon)

- **Name:** Read Aloud
- **Add to Share Sheet:** on
- **Accepted types:** URLs, Files, Rich Text, Text
- **Show in:** Shortcut app, Home Screen, Siri, Apple Watch

---

### Actions

#### 1 — Get URLs from Shortcut Input

| Field | Value |
|-------|-------|
| Action | **Get URLs from** |
| Input | Shortcut Input |

---

#### 2 — If (URL found)

| Field | Value |
|-------|-------|
| Action | **If** |
| Input | *URLs* (result of step 1) |
| Condition | **has any value** |

---

#### 3 — Get Item from List  *(inside the If)*

| Field | Value |
|-------|-------|
| Action | **Get Item from List** |
| Get | **First Item** |

---

#### 4 — Get Contents of URL  *(inside the If)*

| Field | Value |
|-------|-------|
| Action | **Get Contents of URL** |
| Input | (current result — the URL from step 3) |

---

#### 5 — Get Article from Web Page  *(inside the If)*

| Field | Value |
|-------|-------|
| Action | **Get Article from Web Page** |
| Input | (current result — downloaded page from step 4) |

---

#### 6 — Get Detail of Article  *(inside the If)*

| Field | Value |
|-------|-------|
| Action | **Get Detail of Article** |
| Detail | **Body** |

---

#### 7 — Speak Text  *(inside the If)*

| Field | Value |
|-------|-------|
| Action | **Speak Text** |
| Text | (current result — article body from step 6) |
| Rate | 0.45 |
| Pitch | 1.0 |
| Wait until done | on |

---

#### 8 — Otherwise

*(drag this block below the If; all remaining steps until End If go here)*

---

#### 9 — If (no Shortcut Input)  *(inside the Otherwise)*

| Field | Value |
|-------|-------|
| Action | **If** |
| Input | Shortcut Input |
| Condition | **has no value** |

---

#### 10 — Get Clipboard  *(inside the inner If)*

| Field | Value |
|-------|-------|
| Action | **Get Clipboard** |

---

#### 11 — Speak Text  *(inside the inner If)*

| Field | Value |
|-------|-------|
| Action | **Speak Text** |
| Text | (current result — clipboard) |
| Rate | 0.45 |
| Pitch | 1.0 |
| Wait until done | on |

---

#### 12 — Otherwise  *(inner Otherwise — file or text was shared)*

---

#### 13 — Get Text from Input  *(inside the inner Otherwise)*

| Field | Value |
|-------|-------|
| Action | **Get Text from Input** |
| Input | Shortcut Input |

---

#### 14 — Speak Text  *(inside the inner Otherwise)*

| Field | Value |
|-------|-------|
| Action | **Speak Text** |
| Text | (current result — extracted text) |
| Rate | 0.45 |
| Pitch | 1.0 |
| Wait until done | on |

---

#### 15 — End If  *(closes inner If from step 9)*

#### 16 — End If  *(closes outer If from step 2)*

---

## Troubleshooting

**Webpage reads nothing / stops early**
Some pages block automated fetches. Try sharing the page URL directly from
Safari's address bar rather than using the Share Sheet on the page itself.
If the site requires login, the article extractor won't have access — copy the
text manually and run the shortcut standalone to read the clipboard instead.

**File content is not read**
Confirm the file type is supported. PDFs, plain text (.txt), Pages, Word (.docx),
and rich text (.rtf) files work. Images and spreadsheets are not supported.

**Clipboard is read instead of the shared item**
This happens when the Share Sheet sends the item as a URL rather than as a file
or text object. Long-press the Share button in the source app and choose
"Share…" to get the standard Share Sheet with all options.

---

## Regenerating the shortcut file

If you modify `generate_shortcut.py`:

```bash
python3 generate_shortcut.py
# produces ReadAloud.shortcut in the same directory
```

Requires Python 3.6+ (uses only the standard library).
