# Header/Footer Round-Trip Samples

This directory stores semantic samples for header/footer/page-number round-trip checks.

Current CI baseline uses generated fixtures in:
- `structheader_test.go`
- `cmd/rtcheck/main_test.go`

Planned coverage:
- default header + default footer
- all kinds (`default/first/even`)
- footer PAGE field structure
- header/footer containing table and unknown nodes
