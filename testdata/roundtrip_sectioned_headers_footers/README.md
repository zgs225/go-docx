# Sectioned Header/Footer Round-Trip Samples

This directory stores semantic samples for multi-section header/footer/page-number round-trip checks.

Current CI baseline uses generated fixtures in:
- `structheader_test.go`
- `cmd/rtcheck/main_test.go`

Planned coverage:
- 2-3 sections in document order (`pPr/sectPr` + body tail `sectPr`)
- section-scoped header/footer references
- section-scoped page number field and `pgNumType`
- mixed `default/first/even` kinds across sections
