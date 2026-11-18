# Scoped XLSX cell editing

Date: 2026-11-18
Status: Accepted

## Context

The work-panel Office viewer previously exposed spreadsheet values through a
read-only table. Users need a spreadsheet grid with read/edit modes and a
save action while retaining existing workbook content. A full round trip
through a grid exporter can remove charts, drawings, unsupported styles and
other OOXML extensions. Rendering arbitrary external files does not imply
permission to overwrite them.

## Decision

Use the MIT-licensed FortuneSheet React grid for display and value editing.
Keep the renderer responsible for interaction and draft state. Introduce the
main-window-only `fs/saveSpreadsheet` channel, with local host-runtime code
validating and applying explicit cell-value patches to the original XLSX
archive. The existing read-only plugin SDK and remote protocol are unchanged.

The write boundary is narrower than the read boundary: only the effective
local workspace and the owning session's scratch root are eligible. Validate
realpath containment, original SHA-256 revision, patch bounds and values,
workbook protection, signature/macro absence and preview completeness before
writing. Serialize concurrent saves per target. Build and validate the output
first, write a unique sibling temporary file, sync it, recheck containment
and revision, and rename over the source. This reduces partial-file risk;
it is not a cross-process filesystem compare-and-swap primitive.

Preserve original cell style attributes and untouched package parts. Treat
formula cells, shared/array formula sheets and merged follower cells as
read-only. Do not execute formulas or active workbook content. Set workbook
calculation flags so Excel can recalculate on reopening. Keep drafts in
renderer memory across file/session switches and retain them on save errors.

## Alternatives

Continuing the bespoke HTML table would require rebuilding spreadsheet
selection and input behavior. Exporting all grid data as a new workbook
would risk discarding unsupported original parts. A hosted Office editor
would introduce document hosting, authentication and external-service
requirements for local files. FortuneSheet plus bounded OOXML patches fits
the existing local preview pipeline without those changes.

## Consequences

This first version edits text, numbers, booleans and empty cell values, not
formulas, formatting or worksheet structure. Rendering does not promise Excel
visual fidelity. Formula results remain cached until Excel recalculates.
Large/truncated, protected, signed and macro-enabled workbooks remain
read-only. Drafts are not persistent recovery files. Atomic replacement may
fail when another application locks the workbook; errors keep the draft.
