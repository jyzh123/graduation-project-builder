# Skill Maintenance: CAD MTEXT Orientation Audit

Date: 2026-06-04

The SGB620/80T drawing package exposed a false failure in the strict mechanical CAD package audit: standard DXF `MTEXT` entities with group code `71=5` were reported as upside-down and mirrored text. For `MTEXT`, group code `71` is the attachment point, not a text generation flag. Treating it like `TEXT` generation flags caused valid centered notes and dimension text to fail even when rendered review evidence showed legible upright text.

Rule consolidation:

- `scripts/audit_mechanical_drawing_package.py` must parse generation flags from group code `71` only for non-`MTEXT` text entities such as `TEXT`, `ATTRIB`, and `ATTDEF`.
- `MTEXT` orientation must be judged from rotation and direction vector group codes (`11`, `21`) rather than attachment point `71`.
- A self-test fixture with `MTEXT`, `71=5`, and direction vector `(1,0)` must remain a passing case, while the existing inverted `TEXT` fixture must remain a failing case.

This maintenance case hardens the CAD text-orientation gate. It does not weaken the rendered overlap, frame overflow, source linework, color-family, text mojibake, or actual upside-down/mirrored text requirements.
