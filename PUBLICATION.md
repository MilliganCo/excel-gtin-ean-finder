# Public source preview — 2026-10-29

Local Excel barcode analysis utility. Complete reachable history reviewed; no identified credential,
personal-data or private business-record blocker to source publication.

Fixed: pandas now preserves original Excel cell types before barcode normalization. The canonical
entrypoint is analyze_barcodes.py. The old double-extension filename is removed.
PASS — one real synthetic XLSX roundtrip verifies text leading zeros, distinct barcode counts,
numeric auxiliary prices, the problem sheet and unchanged input bytes.
NOT RUN — GUI and legacy XLS. Zeros present only as a numeric cell display format cannot be inferred.
Install requirements.txt; run `python -m unittest discover -s tests -v` to repeat the regression.
