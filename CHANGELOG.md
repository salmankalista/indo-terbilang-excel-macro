# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-11-26

### Added
- Initial release of `indo-terbilang` for Microsoft Excel VBA / Macro (`src/IndoTerbilang.bas`).
- `Terbilang` UDF & VBA function with full EYD V / KBBI support up to 36 digits (`desiliun`), digit-by-digit decimal reading, negative normalization (`-0` -> `"nol"`), and casing styles (`lower`, `upper`, `title`, `sentence`) with `Prefix`/`Suffix`.
- `TerbilangRupiah` UDF & VBA function with exact string-based decimal rounding (`half-up` / `truncate`) and `sen` support.
- `FormatRupiah` UDF & VBA function for deterministic Rupiah formatting and compact Indonesian notation (`rb`, `jt`, `M`, `T`).
- `ParseRupiah` UDF & VBA function to extract canonical numeric strings from formatted Rupiah text (`"Rp 1.500.000,50"`, `"(Rp 500.000)"`).
- `RegisterIndoTerbilangFunctions` helper macro to register descriptions in Excel's Insert Function (`fx`) dialog.
- Cross-platform pure VBA implementation (zero external COM/ActiveX dependencies, compatible with Windows Excel 32/64-bit and Excel for Mac, immune to Windows Regional Settings).
- Example macros (`examples/IndoTerbilangExamples.bas`), native VBA tests (`tests/IndoTerbilangTests.bas`), and automated VBScript + Python parity verification suite (`tests/verify_macro.py`).
