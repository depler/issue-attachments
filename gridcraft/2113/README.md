# storytold/gridcraft#2113 - PivotTable Fields hang

Reproducer for the UI stall when a workbook with a PivotTable is opened: the PivotTable
Fields pane calls `pivot.fields` on every frame, which re-reads the whole pivot source.

- `pivot-fields-hang.xlsx` - 250,000 rows x 20 columns, one PivotTable on `Report`.
  21.0 MB, sha256 `497984af432f15f0681e2ee78a168f59136028cc05608165800993c6bb0b0464`.
- `make_pivot_hang_repro.py` - fully synthetic, seeded generator (no real data):

  ```
  python make_pivot_hang_repro.py 250000 pivot-fields-hang.xlsx
  ```

GridCraft 0.9.0 (Windows x64): `frameMs` 3726 ms, `gridMs` 0.33 ms. Excel opens the file
and rebuilds the report from the source (`saveData="0" refreshOnLoad="1"`).
