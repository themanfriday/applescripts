#!/usr/bin/env python3
"""
Generates ReadAloud.shortcut — an iOS Shortcut that reads aloud:

  • URL  — when a URL is shared (via Share Sheet or Siri), the shortcut
            fetches the webpage, extracts the article body, and speaks it.
  • File — when a file (PDF, .txt, Pages doc, etc.) is shared, the shortcut
            extracts its text content and speaks it.
  • Text — when plain or rich text is shared the shortcut speaks it directly.
  • Clipboard — when run standalone (no Share Sheet input) the shortcut
                reads aloud whatever is on the clipboard.

Usage:
    python3 generate_shortcut.py
    # → produces ReadAloud.shortcut in the same directory

Then AirDrop or copy ReadAloud.shortcut to your iPhone/iPad and open it to
import it into the Shortcuts app.  Alternatively, host it on iCloud Drive and
tap the file from the Files app.
"""

import plistlib
import uuid
from pathlib import Path

# ── helpers ──────────────────────────────────────────────────────────────────

def uid() -> str:
    return str(uuid.uuid4()).upper()


def shortcut_input_ref() -> dict:
    """Magic variable: the item(s) passed in via the Share Sheet / Siri."""
    return {
        "Value": {"Type": "ExtensionInput"},
        "WFSerializationType": "WFTextTokenAttachment",
    }


def output_ref(name: str, action_uuid: str) -> dict:
    """Reference the output of a specific action by UUID."""
    return {
        "Value": {
            "OutputName": name,
            "OutputUUID": action_uuid,
            "Type": "ActionOutput",
        },
        "WFSerializationType": "WFTextTokenAttachment",
    }


def action(identifier: str, **params) -> dict:
    return {
        "WFWorkflowActionIdentifier": identifier,
        "WFWorkflowActionParameters": {"UUID": uid(), **params},
    }


def if_cond(condition: int, wf_input: dict) -> dict:
    return {
        "WFWorkflowActionIdentifier": "is.workflow.actions.conditional",
        "WFWorkflowActionParameters": {
            "UUID": uid(),
            "WFControlFlowMode": 0,   # 0 = If
            "WFCondition": condition,
            "WFInput": wf_input,
        },
    }


def otherwise() -> dict:
    return {
        "WFWorkflowActionIdentifier": "is.workflow.actions.conditional",
        "WFWorkflowActionParameters": {"UUID": uid(), "WFControlFlowMode": 1},
    }


def endif() -> dict:
    return {
        "WFWorkflowActionIdentifier": "is.workflow.actions.conditional",
        "WFWorkflowActionParameters": {"UUID": uid(), "WFControlFlowMode": 2},
    }


# ── constants ─────────────────────────────────────────────────────────────────

HAS_ANY_VALUE = 8000   # "has any value" condition
HAS_NO_VALUE  = 8001   # "has no value"  condition

# Speak Text settings applied consistently across all branches
SPEAK = {
    "WFSpeakTextRate": 0.45,       # slightly slower than default for clarity
    "WFSpeakTextPitch": 1.0,
    "WFSpeakTextWait": True,        # block until speech finishes
    "WFSpeakTextLanguageCode": "",  # empty → use device language
    "ShowWhenRun": False,
}

# ── action list ───────────────────────────────────────────────────────────────
#
# Control flow overview:
#
#   if (Shortcut Input contains a URL):
#     fetch webpage → extract article body → speak
#   else:
#     if (Shortcut Input is empty, i.e. run standalone):
#       get clipboard → speak
#     else:                                   ← file or plain text was shared
#       get text from input → speak
#     end if
#   end if

detect_link_uuid = uid()   # fixed so we can reference its output in the If

actions = [

    # ── 1. Detect URLs inside whatever was passed in as Shortcut Input ──────
    {
        "WFWorkflowActionIdentifier": "is.workflow.actions.detect.link",
        "WFWorkflowActionParameters": {
            "UUID": detect_link_uuid,
            "WFInput": shortcut_input_ref(),
        },
    },

    # ── 2. If at least one URL was found → webpage path ─────────────────────
    if_cond(HAS_ANY_VALUE, output_ref("Links", detect_link_uuid)),

    #   2a. Take the first URL from the detected list
    action("is.workflow.actions.getitemfromlist", WFItemSpecifier="First Item"),

    #   2b. Download the webpage HTML
    action("is.workflow.actions.downloadurl"),

    #   2c. Parse the downloaded content into a structured article
    #       (equivalent to Safari Reader — extracts title, body, author, etc.)
    action("is.workflow.actions.getarticle"),

    #   2d. Pull just the body text from the article
    action(
        "is.workflow.actions.article.getdetail",
        WFArticleDetailName="Body",
    ),

    #   2e. Speak the article body
    action("is.workflow.actions.speaktext", **SPEAK),

    # ── 3. Otherwise → input is not a URL (or no URL detected) ──────────────
    otherwise(),

    # ── 3a. If Shortcut Input has no value → standalone / clipboard path ─────
    if_cond(HAS_NO_VALUE, shortcut_input_ref()),

    #   Get whatever is on the clipboard and speak it as-is
    action("is.workflow.actions.getclipboard"),
    action("is.workflow.actions.speaktext", **SPEAK),

    # ── 3b. Otherwise → a file or text item was shared ──────────────────────
    otherwise(),

    #   Extract plain text from the shared item.
    #   Works with: .txt, PDF, Pages/Word docs, rich text, plain strings.
    {
        "WFWorkflowActionIdentifier": "is.workflow.actions.detect.text",
        "WFWorkflowActionParameters": {
            "UUID": uid(),
            "WFInput": shortcut_input_ref(),
        },
    },
    action("is.workflow.actions.speaktext", **SPEAK),

    endif(),   # close inner  if (no input)
    endif(),   # close outer  if (URL detected)
]

# ── shortcut metadata ─────────────────────────────────────────────────────────

shortcut = {
    # Client version string from a recent iOS export; safe to leave as-is.
    "WFWorkflowClientVersion": "1284.2",

    "WFWorkflowHasOutputFallback": False,
    "WFWorkflowHasShortcutInputVariables": True,

    # Icon: speaker glyph (59511), purple background
    "WFWorkflowIcon": {
        "WFWorkflowIconGlyphNumber": 59511,
        "WFWorkflowIconStartColor": 4071018751,
    },

    "WFWorkflowImportQuestions": [],

    # Content types accepted in the Share Sheet
    "WFWorkflowInputContentItemClasses": [
        "WFArticleContentItem",
        "WFGenericFileContentItem",
        "WFPDFContentItem",
        "WFRichTextContentItem",
        "WFSafariWebPageContentItem",
        "WFStringContentItem",
        "WFURLContentItem",
    ],

    "WFWorkflowMinimumClientVersion": 900,
    "WFWorkflowMinimumClientVersionString": "900",
    "WFWorkflowOutputContentItemClasses": [],

    # Where the shortcut appears:
    #   ActionExtension → Share Sheet
    #   NCWidget        → Home Screen widget
    #   WatchKit        → Apple Watch
    "WFWorkflowTypes": ["ActionExtension", "NCWidget", "WatchKit"],

    "WFQuickActionSurfaces": [],
    "WFWorkflowActions": actions,
}

# ── write binary plist ────────────────────────────────────────────────────────

output_path = Path(__file__).parent / "ReadAloud.shortcut"
with open(output_path, "wb") as fh:
    plistlib.dump(shortcut, fh, fmt=plistlib.FMT_BINARY)

print(f"Generated: {output_path}  ({output_path.stat().st_size} bytes)")
print(f"Actions:   {len(actions)}")
