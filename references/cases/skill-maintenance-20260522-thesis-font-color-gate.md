# 2026-05-22 Thesis Font-Color Gate

## Trigger

The user reported that a generated graduation thesis showed blue visible text and that later repair passes did not restore the text to black.

## Root Cause

The manuscript did not carry blue direct run colors in body text. The visible blue color came from used Word styles such as `Heading1`, `Heading2`, `Heading3`, `Title`, and `Caption`, whose style definitions retained `accent1` or `text2` theme colors in `word/styles.xml` and `word/stylesWithEffects.xml`. Previous checks counted references, figures, PDF export, and direct text content but did not fail on used style color drift.

## Durable Change

- Added `FB-LAYOUT-073` in `references/user-feedback/template-and-layout.md`.
- Added `scripts/audit_docx_font_color.py` to audit direct run colors plus actually used style colors and optionally write a repaired DOCX with black visible font color.
- Added final acceptance fields for exact-output font-color audit path and verdict.
- Registered the rule in `references/rule-owner-map.json` with `scripts/audit_docx_font_color.py` as validator owner.

## Validation Scope

The current project repair must rerun the font-color audit on the exact final DOCX, regenerate the PDF from that DOCX, and rebuild the final handoff package. Skill validation must include Python syntax, rule-owner JSON parse, UTF-8 clean, and `validate_skill_gate.py`.
