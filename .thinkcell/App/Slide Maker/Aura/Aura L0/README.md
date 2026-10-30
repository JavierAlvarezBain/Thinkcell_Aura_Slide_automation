# Aura L0 module map

Build identifier: `l0-dynamic-ten-shades-lightest-overflow-active-stream-v2026-07-31`

The accepted July 31 L0 implementation is split by responsibility. L0 does not
reuse the Aura L1 quartile or color handlers.

- `l0_config.py` — paths, template contract, constants, and data models.
- `l0_source.py` — shared Aura-table reading plus L0-only grouping, direct-L0
  quartile detection, peer percentile fallback, client classification, segment
  ordering, Bain family color generation, and ChartData workbook creation.
- `l0_powerpoint.py` — L0 title/agenda safety, section-deck collation, live chart
  series formatting helpers, and visible segment legend handling.
- `l0_native.py` — L0-specific value-line identities, compound-file access,
  active `think-cellXML` colors/legend/quartile patching, and selected-slide
  final-deck patching.
- `l0_builder.py` — the complete chart-creation loop. It can either finalize an
  L0-only deck or return an unpatched L0 section plus the metadata needed for a
  later final-deck patch.
- `l0_manifest.py` — compact, serializable slide records used to reproduce the
  exact L0 native patch during final-merge recovery without recalculating L0.
- `run_aura_l0.py` — standalone CLI runner.
- `L0_test.py` — compatibility launcher that calls `run_aura_l0.py`.

## Consolidated workflow

`python app.py` now builds sections in this order:

1. Aura Summary
2. Aura L0
3. Aura L1
4. Global Geo Cost, when detected/requested
5. US Geo Cost, when detected/requested

The L0 section is deliberately merged as a shell. After the isolated PowerPoint
merge process has saved and closed the consolidated presentation,
`patch_l0_slides_in_deck(...)` locates only the registered L0 slide embeddings
and writes native segment colors, native legend order, and L0 quartile geometry.
No PowerPoint reopen or save is permitted after that patch.

Use `--skip-l0` to omit L0 from the consolidated workflow. Use
`--skip-l0-quartiles` to preserve the unpatched L0 section-shell behavior. Use
`python app.py aura-l0 ...` for an L0-only deck.

Live PowerPoint/Excel/think-cell execution remains the runtime acceptance test.
