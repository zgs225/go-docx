# Round-Trip Unknown-Node Samples

This directory contains round-trip golden samples focused on unknown-node preservation.

Planned coverage:
- paragraphs containing fields/TOC-related markup
- body-level unknown/extension nodes (for preservation checks)
- mixed table content with nested table cells and unknown nodes
- runs containing drawing/shape and unsupported siblings

Current CI baseline uses generated fixtures in:
- `unknown_node_roundtrip_test.go`
- `cmd/rtcheck/main_test.go`

You can run the structural checker locally with:

```bash
go run ./cmd/rtcheck --in ./testdata/roundtrip_unknown_nodes --out /tmp/rtcheck-out --report /tmp/rtcheck-report.json
```
