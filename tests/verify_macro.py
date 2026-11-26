"""Automated static checker, CRLF normalizer check, VBScript test runner, and parity oracle."""

from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = ROOT.parent
BAS_MAIN = ROOT / "src" / "IndoTerbilang.bas"
BAS_EXAMPLES = ROOT / "examples" / "IndoTerbilangExamples.bas"
BAS_TESTS = ROOT / "tests" / "IndoTerbilangTests.bas"

PYTHON_SRC = REPO_ROOT / "code-python" / "src"
HAS_PYTHON_ORACLE = (PYTHON_SRC / "indo_terbilang").is_dir()
if HAS_PYTHON_ORACLE:
    sys.path.insert(0, str(PYTHON_SRC))
    from indo_terbilang import (  # type: ignore[import-not-found]
        TerbilangError as OracleError,
        format_rupiah as oracle_format_rupiah,
        parse_rupiah as oracle_parse_rupiah,
        terbilang as oracle_terbilang,
        terbilang_rupiah as oracle_terbilang_rupiah,
    )
else:
    OracleError = None  # type: ignore[assignment,misc]
    oracle_format_rupiah = None
    oracle_parse_rupiah = None
    oracle_terbilang = None
    oracle_terbilang_rupiah = None


def ensure_crlf_and_check_bas(path: Path, expected_vb_name: str) -> str:
    raw_bytes = path.read_bytes()
    # Ensure pure 7-bit ASCII so Excel VBA VBE import never mojibakes across Windows codepages.
    try:
        text = raw_bytes.decode("ascii")
    except UnicodeDecodeError as exc:
        raise AssertionError(f"{path.name} contains non-ASCII byte at {exc.start}") from exc

    # Normalize to CRLF if written with LF by an editor/tool
    if b"\n" in raw_bytes and b"\r\n" not in raw_bytes:
        normalized = text.replace("\n", "\r\n").encode("ascii")
        path.write_bytes(normalized)
        raw_bytes = normalized

    lines = text.splitlines()
    if not lines or lines[0] != f'Attribute VB_Name = "{expected_vb_name}"':
        raise AssertionError(f"{path.name} missing Attribute VB_Name = \"{expected_vb_name}\"")
    if "Option Explicit" not in lines[:5]:
        raise AssertionError(f"{path.name} missing Option Explicit near top")

    banned_tokens = ("CreateObject", "VBScript.RegExp", "Declare PtrSafe", "Declare Function")
    for token in banned_tokens:
        if token in text:
            raise AssertionError(f"{path.name} contains banned external dependency token: {token}")

    if "Replace$" in text:
        raise AssertionError(f"{path.name} uses invalid VBA function Replace$ (use Replace)")

    return text


def transpile_bas_to_vbs(vba_source: str) -> str:
    """Mechanically adapt IndoTerbilang.bas to run inside Windows cscript.exe VBScript."""
    out_lines: list[str] = [
        "Option Explicit",
        "",
        "Function IsError(v)",
        "    IsError = (VarType(v) = vbError)",
        "End Function",
        "",
        "Function VbaStr(v)",
        "    Dim s",
        "    s = CStr(v)",
        "    If InStr(1, s, \",\", 0) > 0 Then s = Replace(s, \",\", \".\")",
        "    ' Simulate VBA Str$ omitting leading zero for (-1, 1) numbers",
        "    If Left(s, 2) = \"0.\" Then s = Mid(s, 2)",
        "    If Left(s, 3) = \"-0.\" Then s = \"-\" & Mid(s, 3)",
        "    If Left(s, 1) <> \"-\" Then s = \" \" & s",
        "    VbaStr = s",
        "End Function",
        "",
    ]

    lines = vba_source.splitlines()
    skip_until_end_block = False
    end_block_token = ""

    # Join line continuations (_) first so signatures are single-line
    joined: list[str] = []
    buf = ""
    for line in lines:
        stripped = line.rstrip()
        if stripped.endswith(" _"):
            buf += stripped[:-1]
        else:
            joined.append(buf + line)
            buf = ""
    if buf:
        joined.append(buf)

    for line in joined:
        trimmed = line.strip()

        if skip_until_end_block:
            if trimmed == end_block_token:
                skip_until_end_block = False
            continue

        if trimmed.startswith("Attribute VB_Name") or trimmed == "Option Explicit":
            continue

        # Skip VBA-only Excel wrapper functions/subs (we test *Core + error behavior directly)
        if re.match(
            r"^Public (?:Function (?:Terbilang|TerbilangRupiah|FormatRupiah|ParseRupiah)|"
            r"Sub RegisterIndoTerbilangFunctions)\b",
            trimmed,
        ):
            skip_until_end_block = True
            end_block_token = "End Function" if "Function " in trimmed else "End Sub"
            continue

        if re.match(r"^Private (?:Sub TryRegisterMacroOption|Function IsWorksheetCall|Function MapErrToExcelError)\b", trimmed):
            skip_until_end_block = True
            end_block_token = "End Function" if "Function " in trimmed else "End Sub"
            continue

        # Strip comments safely only after checking structure
        if trimmed.startswith("'"):
            out_lines.append(line)
            continue

        # Convert Public/Private Const X As Type = expr -> Const X = literal
        m_const = re.match(r"^(?:Public|Private)\s+Const\s+(\w+)\s+As\s+\w+\s*=\s*(.+)$", trimmed)
        if m_const:
            const_name, const_expr = m_const.group(1), m_const.group(2).strip()
            m_vbo = re.match(r"^vbObjectError\s*\+\s*(\d+)$", const_expr)
            if m_vbo:
                const_expr = str(-2147221504 + int(m_vbo.group(1)))
            out_lines.append(f"Const {const_name} = {const_expr}")
            continue

        # Convert Dim var() As Type -> Dim var
        line = re.sub(r"\bDim\s+(\w+)\s*\(\s*\)\s+As\s+\w+", r"Dim \1", line)
        # Convert Dim arr(0 To N) As Type -> Dim arr(N)
        line = re.sub(r"\bDim\s+(\w+)\s*\(\s*0\s+To\s+(\d+)\s*\)\s+As\s+\w+", r"Dim \1(\2)", line)
        # Convert Dim x As Type -> Dim x
        line = re.sub(r"\bDim\s+(\w+)\s+As\s+\w+", r"Dim \1", line)

        # Convert Private Function/Sub signatures (strip ByVal/ByRef/As Type)
        if re.match(r"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+", line):
            line = re.sub(r"^\s*(?:Public|Private)\s+", "", line)
            line = re.sub(r"\b(?:ByVal|ByRef)\s+", "", line)
            line = re.sub(r"\s+As\s+\w+(?=\s*[,)])", "", line)
            line = re.sub(r"\)\s+As\s+\w+\s*$", ")", line)

        # Replace VBA $ string suffixes: Mid$, Left$, Right$, Trim$, LCase$, UCase$, String$, Str$
        line = re.sub(r"\bStr\$\s*\(", "VbaStr(", line)
        line = re.sub(r"\b(Mid|Left|Right|Trim|LCase|UCase|String)\$", r"\1", line)

        # Replace Mid(s, i, 1) = ch statement with VBS substring splice
        m_mid_assign = re.match(r"^(\s*)Mid\((\w+),\s*([^,]+),\s*1\)\s*=\s*(.+)$", line)
        if m_mid_assign:
            indent, var_name, idx_expr, val_expr = m_mid_assign.groups()
            line = (
                f"{indent}{var_name} = Left({var_name}, ({idx_expr}) - 1) "
                f"& ({val_expr}) & Mid({var_name}, ({idx_expr}) + 1)"
            )

        # VBScript does not allow variable name after Next
        line = re.sub(r"^(\s*Next)\s+\w+\s*$", r"\1", line)

        out_lines.append(line)

    return "\r\n".join(out_lines) + "\r\n"


def vbs_quote(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def build_parity_cases() -> list[tuple[str, str, str]]:
    """Return (label, vbs_expression, expected_output), cross-checked against code-python if present."""
    cases: list[tuple[str, str, str]] = []

    terbilang_inputs: list[tuple[str, object, str, str, str, str, str, bool, str]] = [
        ("zero int", 0, "lower", "", "", "minus", "koma", False, "nol"),
        ("neg zero str", "-0.00", "lower", "", "", "minus", "koma", True, "nol"),
        ("1", 1, "lower", "", "", "minus", "koma", False, "satu"),
        ("10", 10, "lower", "", "", "minus", "koma", False, "sepuluh"),
        ("11", 11, "lower", "", "", "minus", "koma", False, "sebelas"),
        ("19", 19, "lower", "", "", "minus", "koma", False, "sembilan belas"),
        ("21", 21, "lower", "", "", "minus", "koma", False, "dua puluh satu"),
        ("100", 100, "lower", "", "", "minus", "koma", False, "seratus"),
        ("111", 111, "lower", "", "", "minus", "koma", False, "seratus sebelas"),
        ("1000", 1000, "lower", "", "", "minus", "koma", False, "seribu"),
        ("1001", 1001, "lower", "", "", "minus", "koma", False, "seribu satu"),
        ("21000", 21000, "lower", "", "", "minus", "koma", False, "dua puluh satu ribu"),
        ("101000", 101000, "lower", "", "", "minus", "koma", False, "seratus satu ribu"),
        ("1000000", 1000000, "lower", "", "", "minus", "koma", False, "satu juta"),
        ("1001000", 1001000, "lower", "", "", "minus", "koma", False, "satu juta seribu"),
        ("1e9", 1000000000, "lower", "", "", "minus", "koma", False, "satu miliar"),
        ("1e12", 1000000000000, "lower", "", "", "minus", "koma", False, "satu triliun"),
        ("1e21 sci", "1e21", "lower", "", "", "minus", "koma", False, "satu sekstiliun"),
        ("1e-7 sci", "1e-7", "lower", "", "", "minus", "koma", False, "nol koma nol nol nol nol nol nol satu"),
        ("36 digits", "1" + "0" * 35, "lower", "", "", "minus", "koma", False, "seratus desiliun"),
        (
            "999 desiliun",
            "9" * 36,
            "lower",
            "",
            "",
            "minus",
            "koma",
            False,
            "sembilan ratus sembilan puluh sembilan desiliun "
            "sembilan ratus sembilan puluh sembilan noniliun "
            "sembilan ratus sembilan puluh sembilan oktiliun "
            "sembilan ratus sembilan puluh sembilan septiliun "
            "sembilan ratus sembilan puluh sembilan sekstiliun "
            "sembilan ratus sembilan puluh sembilan kuintiliun "
            "sembilan ratus sembilan puluh sembilan kuadriliun "
            "sembilan ratus sembilan puluh sembilan triliun "
            "sembilan ratus sembilan puluh sembilan miliar "
            "sembilan ratus sembilan puluh sembilan juta "
            "sembilan ratus sembilan puluh sembilan ribu "
            "sembilan ratus sembilan puluh sembilan",
        ),
        ("decimal 3.14", 3.14, "lower", "", "", "minus", "koma", False, "tiga koma satu empat"),
        ("decimal 12.05", "12.05", "lower", "", "", "minus", "koma", False, "dua belas koma nol lima"),
        ("decimal 12.50", "12.50", "lower", "", "", "minus", "koma", False, "dua belas koma lima nol"),
        ("decimal 12.50 trim", "12.50", "lower", "", "", "minus", "koma", True, "dua belas koma lima"),
        ("decimal 12.00 trim", "12.00", "lower", "", "", "minus", "koma", True, "dua belas"),
        ("comma decimal .5", ",5", "lower", "", "", "minus", "koma", False, "nol koma lima"),
        ("negative -1500", -1500, "lower", "", "", "minus", "koma", False, "minus seribu lima ratus"),
        ("style title affixes", 1500000, "title", "Terbilang:", "rupiah", "minus", "koma", False, "Terbilang: Satu Juta Lima Ratus Ribu Rupiah"),
        ("style upper", 25, "upper", "", "", "minus", "koma", False, "DUA PULUH LIMA"),
        ("style sentence", 25, "sentence", "", "", "minus", "koma", False, "Dua puluh lima"),
    ]

    for label, val, style, pref, suff, neg_p, dec_w, trim_z, expected in terbilang_inputs:
        if oracle_terbilang is not None:
            oracle_out = oracle_terbilang(
                val,  # type: ignore[arg-type]
                style=style,  # type: ignore[arg-type]
                prefix=pref or None,
                suffix=suff or None,
                negative_prefix=neg_p,
                decimal_word=dec_w,
                trim_trailing_decimal_zeros=trim_z,
            )
            assert oracle_out == expected, f"Oracle mismatch for {label}: {oracle_out!r} != {expected!r}"
        vba_val = vbs_quote(val) if isinstance(val, str) else str(val)
        vbs_expr = (
            f"TerbilangCore({vba_val}, {vbs_quote(style)}, {vbs_quote(pref)}, "
            f"{vbs_quote(suff)}, {vbs_quote(neg_p)}, {vbs_quote(dec_w)}, {str(trim_z)})"
        )
        cases.append((f"Terbilang:{label}", vbs_expr, expected))

    rupiah_inputs: list[tuple[str, object, str, str, str, str, str, bool, str, str, str]] = [
        ("0", 0, "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "nol rupiah"),
        ("1250.5", 1250.5, "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "seribu dua ratus lima puluh rupiah lima puluh sen"),
        ("1250.05", "1250.05", "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "seribu dua ratus lima puluh rupiah lima sen"),
        ("1250.00", "1250.00", "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "seribu dua ratus lima puluh rupiah"),
        ("1250.567", "1250.567", "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "seribu dua ratus lima puluh rupiah lima puluh tujuh sen"),
        ("1250.996", "1250.996", "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "seribu dua ratus lima puluh satu rupiah"),
        ("1250.996 trunc", "1250.996", "lower", "", "", "minus", "rupiah", True, "sen", "truncate", "seribu dua ratus lima puluh rupiah sembilan puluh sembilan sen"),
        ("-0.001", "-0.001", "lower", "", "", "minus", "rupiah", True, "sen", "half-up", "nol rupiah"),
        ("1250.6 no sen", 1250.6, "sentence", "", "", "minus", "rupiah", False, "sen", "half-up", "Seribu dua ratus lima puluh satu rupiah"),
        ("1250750.50 title", "1250750.50", "title", "", "", "minus", "rupiah", True, "sen", "half-up", "Satu Juta Dua Ratus Lima Puluh Ribu Tujuh Ratus Lima Puluh Rupiah Lima Puluh Sen"),
        ("custom unit", "2500.25", "title", "Jumlah:", "sah", "minus", "IDR", True, "Sen", "half-up", "Jumlah: Dua Ribu Lima Ratus Idr Dua Puluh Lima Sen Sah"),
    ]

    for label, val, style, pref, suff, neg_p, c_unit, inc_sen, s_unit, rnd, expected in rupiah_inputs:
        if oracle_terbilang_rupiah is not None:
            oracle_out = oracle_terbilang_rupiah(
                val,  # type: ignore[arg-type]
                style=style,  # type: ignore[arg-type]
                prefix=pref or None,
                suffix=suff or None,
                negative_prefix=neg_p,
                currency_unit=c_unit,
                include_sen=inc_sen,
                sen_unit=s_unit,
                rounding=rnd,  # type: ignore[arg-type]
            )
            assert oracle_out == expected, f"Oracle mismatch for {label}: {oracle_out!r} != {expected!r}"
        vba_val = vbs_quote(val) if isinstance(val, str) else str(val)
        vbs_expr = (
            f"TerbilangRupiahCore({vba_val}, {vbs_quote(style)}, {vbs_quote(pref)}, "
            f"{vbs_quote(suff)}, {vbs_quote(neg_p)}, {vbs_quote(c_unit)}, {str(inc_sen)}, "
            f"{vbs_quote(s_unit)}, {vbs_quote(rnd)})"
        )
        cases.append((f"TerbilangRupiah:{label}", vbs_expr, expected))

    format_inputs: list[tuple[str, object, str, bool, str, str, object, str, str, bool, int, str]] = [
        ("0", 0, "Rp", True, ".", ",", 2, "half-up", "minus", False, 1, "Rp 0,00"),
        ("1500000", 1500000, "Rp", True, ".", ",", 2, "half-up", "minus", False, 1, "Rp 1.500.000,00"),
        ("1250.567", "1250.567", "Rp", True, ".", ",", 2, "half-up", "minus", False, 1, "Rp 1.250,57"),
        ("0 dec", 1500000, "Rp", True, ".", ",", 0, "half-up", "minus", False, 1, "Rp 1.500.000"),
        ("auto dec", "1500.50", "Rp", True, ".", ",", "auto", "half-up", "minus", False, 1, "Rp 1.500,5"),
        ("custom sep", 1500.25, "IDR", False, ",", ".", 2, "half-up", "minus", False, 1, "IDR1,500.25"),
        ("parentheses", -1500000, "Rp", True, ".", ",", 2, "half-up", "parentheses", False, 1, "(Rp 1.500.000,00)"),
        ("-0.001", "-0.001", "Rp", True, ".", ",", 2, "half-up", "parentheses", False, 1, "Rp 0,00"),
        ("compact 500", 500, "Rp", True, ".", ",", 2, "half-up", "minus", True, 1, "Rp 500"),
        ("compact 1.5 jt", 1500000, "Rp", True, ".", ",", 2, "half-up", "minus", True, 1, "Rp 1,5 jt"),
        ("compact 2.35 M", 2350000000, "Rp", True, ".", ",", 2, "half-up", "minus", True, 2, "Rp 2,35 M"),
        ("compact promote rb", "999.96", "Rp", True, ".", ",", 2, "half-up", "minus", True, 1, "Rp 1 rb"),
        ("compact promote jt", 999990, "Rp", True, ".", ",", 2, "half-up", "minus", True, 1, "Rp 1 jt"),
        ("compact 1500 T", "1500000000000000", "Rp", True, ".", ",", 2, "half-up", "minus", True, 1, "Rp 1.500 T"),
    ]

    for label, val, sym, space_s, th_sep, dec_sep, decs, rnd, neg_f, cmp_f, cmp_d, expected in format_inputs:
        if oracle_format_rupiah is not None:
            oracle_out = oracle_format_rupiah(
                val,  # type: ignore[arg-type]
                symbol=sym,
                space_after_symbol=space_s,
                thousands_separator=th_sep,
                decimal_separator=dec_sep,
                decimals=decs,  # type: ignore[arg-type]
                rounding=rnd,  # type: ignore[arg-type]
                negative_format=neg_f,  # type: ignore[arg-type]
                compact=cmp_f,
                compact_decimals=cmp_d,
            )
            assert oracle_out == expected, f"Oracle mismatch for {label}: {oracle_out!r} != {expected!r}"
        vba_val = vbs_quote(val) if isinstance(val, str) else str(val)
        vba_decs = vbs_quote(decs) if isinstance(decs, str) else str(decs)
        vbs_expr = (
            f"FormatRupiahCore({vba_val}, {vbs_quote(sym)}, {str(space_s)}, "
            f"{vbs_quote(th_sep)}, {vbs_quote(dec_sep)}, {vba_decs}, {vbs_quote(rnd)}, "
            f"{vbs_quote(neg_f)}, {str(cmp_f)}, {cmp_d})"
        )
        cases.append((f"FormatRupiah:{label}", vbs_expr, expected))

    parse_inputs = [
        ("Rp 1.500.000,50", "1500000.50"),
        ("Rp1.500.000", "1500000"),
        ("IDR 2.500.000,75", "2500000.75"),
        ("1.500.000", "1500000"),
        ("+Rp 500.000", "500000"),
        ("-Rp 500.000,25", "-500000.25"),
        ("(Rp 500.000,25)", "-500000.25"),
        ("(Rp 0,00)", "0.00"),
        ("Rp 001.500", "1500"),
    ]
    for raw, expected in parse_inputs:
        if oracle_parse_rupiah is not None:
            oracle_out = oracle_parse_rupiah(raw)
            assert oracle_out == expected, f"Oracle mismatch for {raw}: {oracle_out!r} != {expected!r}"
        cases.append((f"ParseRupiah:{raw}", f"ParseRupiahCore({vbs_quote(raw)})", expected))

    return cases


def build_error_cases() -> list[tuple[str, str, str]]:
    """Return (label, vbs_stmt, expected_err_const), cross-checked against code-python if present."""
    err_map = {
        "INVALID_INPUT": "TERBILANG_ERR_INVALID_INPUT",
        "NON_FINITE_NUMBER": "TERBILANG_ERR_NON_FINITE",
        "OUT_OF_RANGE": "TERBILANG_ERR_OUT_OF_RANGE",
    }

    cases: list[tuple[str, str, str]] = []
    bad_terbilang: list[tuple[str, object, str]] = [
        ("empty", "", "INVALID_INPUT"),
        ("whitespace", "   ", "INVALID_INPUT"),
        ("internal newline", "123\n4", "INVALID_INPUT"),
        ("exp newline", "1e\n2", "INVALID_INPUT"),
        ("bool true", True, "INVALID_INPUT"),
        ("dot group", "1.500.000", "INVALID_INPUT"),
        ("mixed sep", "1.500,00", "INVALID_INPUT"),
        ("trailing dot", "12.", "INVALID_INPUT"),
        ("37 digits", "1" + "0" * 36, "OUT_OF_RANGE"),
        ("exp 101", "1e101", "OUT_OF_RANGE"),
        ("exp huge overflow guard", "1e99999999999", "OUT_OF_RANGE"),
    ]
    for label, val, expected_code in bad_terbilang:
        if oracle_terbilang is not None and OracleError is not None:
            try:
                oracle_terbilang(val)  # type: ignore[arg-type]
            except OracleError as exc:
                assert exc.code == expected_code, f"Oracle code mismatch for {label}: {exc.code} != {expected_code}"
            else:
                raise AssertionError(f"Expected error for {label}")
        if isinstance(val, bool):
            vba_val = "True"
        else:
            vba_val = vbs_quote(str(val)).replace("\n", '" & vbLf & "')
        stmt = f'res = TerbilangCore({vba_val}, "lower", "", "", "minus", "koma", False)'
        cases.append((f"Err:Terbilang:{label}", stmt, err_map[expected_code]))

    bad_parse = [
        ("", "INVALID_INPUT"),
        ("Rp ", "INVALID_INPUT"),
        ("Rp abc", "INVALID_INPUT"),
        ("Rp 1.50.000", "INVALID_INPUT"),
        ("Rp 1..000", "INVALID_INPUT"),
        ("(-Rp 500)", "INVALID_INPUT"),
        ("Rp " + "1" * 37, "OUT_OF_RANGE"),
    ]
    for raw, expected_code in bad_parse:
        if oracle_parse_rupiah is not None and OracleError is not None:
            try:
                oracle_parse_rupiah(raw)
            except OracleError as exc:
                assert exc.code == expected_code, f"Oracle code mismatch for {raw}: {exc.code} != {expected_code}"
            else:
                raise AssertionError(f"Expected error for {raw}")
        stmt = f"res = ParseRupiahCore({vbs_quote(raw)})"
        cases.append((f"Err:ParseRupiah:{raw}", stmt, err_map[expected_code]))

    return cases


def main() -> None:
    src_text = ensure_crlf_and_check_bas(BAS_MAIN, "IndoTerbilang")
    ensure_crlf_and_check_bas(BAS_EXAMPLES, "IndoTerbilangExamples")
    ensure_crlf_and_check_bas(BAS_TESTS, "IndoTerbilangTests")

    vbs_lib = transpile_bas_to_vbs(src_text)
    parity_cases = build_parity_cases()
    error_cases = build_error_cases()

    test_lines = [
        vbs_lib,
        "Dim passed, failed, res",
        "passed = 0",
        "failed = 0",
        "",
        "Sub CheckEq(expected, actual, label)",
        "    If CStr(actual) = CStr(expected) Then",
        "        passed = passed + 1",
        "    Else",
        "        failed = failed + 1",
        '        WScript.Echo "FAIL [" & label & "] Expected=" & expected & " Actual=" & actual',
        "    End If",
        "End Sub",
        "",
        "Sub CheckErr(expectedNum, actualNum, label)",
        "    If actualNum = expectedNum Then",
        "        passed = passed + 1",
        "    Else",
        "        failed = failed + 1",
        '        WScript.Echo "FAIL [" & label & "] ExpectedErr=" & expectedNum & " ActualErr=" & actualNum',
        "    End If",
        "End Sub",
        "",
    ]

    for label, vbs_expr, expected in parity_cases:
        test_lines.append(f"CheckEq {vbs_quote(expected)}, {vbs_expr}, {vbs_quote(label)}")

    # Extra VBA-specific Str$ edge cases (e.g. 0.5, -0.5, Currency, double 1e-7)
    test_lines.extend([
        'CheckEq "nol koma lima", TerbilangCore(0.5, "lower", "", "", "minus", "koma", False), "VBA:double 0.5"',
        'CheckEq "minus nol koma lima", TerbilangCore(-0.5, "lower", "", "", "minus", "koma", False), "VBA:double -0.5"',
        'CheckEq "nol koma nol nol nol nol nol nol satu", TerbilangCore(1E-07, "lower", "", "", "minus", "koma", False), "VBA:double 1e-7"',
        'CheckEq "seribu dua ratus tiga puluh empat", TerbilangCore(CCur(1234), "lower", "", "", "minus", "koma", False), "VBA:Currency 1234"',
        'CheckEq "seribu dua ratus tiga puluh empat rupiah lima puluh enam sen", TerbilangRupiahCore(CCur(1234.56), "lower", "", "", "minus", "rupiah", True, "sen", "half-up"), "VBA:Currency 1234.56"',
        'CheckEq "seribu rupiah", TerbilangRupiahCore("  1000" & vbLf, "lower", "", "", "minus", "rupiah", True, "sen", "half-up"), "VBA:outer LF"',
        'CheckEq "seribu rupiah", TerbilangRupiahCore(ChrW(160) & "1000" & ChrW(160), "lower", "", "", "minus", "rupiah", True, "sen", "half-up"), "VBA:NBSP"',
    ])

    test_lines.append("On Error Resume Next")
    for label, stmt, expected_err in error_cases:
        test_lines.append("Err.Clear")
        test_lines.append(stmt)
        test_lines.append(f"CheckErr {expected_err}, Err.Number, {vbs_quote(label)}")

    # Additional option & VBA type validation error cases
    test_lines.extend([
        "Err.Clear",
        'res = TerbilangCore(Empty, "lower", "", "", "minus", "koma", False)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:Empty"',
        "Err.Clear",
        'res = TerbilangCore(Null, "lower", "", "", "minus", "koma", False)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:Null"',
        "Err.Clear",
        'res = TerbilangCore(DateSerial(2026, 11, 26), "lower", "", "", "minus", "koma", False)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:Date"',
        "Err.Clear",
        'res = TerbilangCore(Array(1, 2), "lower", "", "", "minus", "koma", False)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:Array"',
        "Err.Clear",
        'res = TerbilangCore(Nothing, "lower", "", "", "minus", "koma", False)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:Nothing"',
        "Err.Clear",
        'res = TerbilangCore(10, "camel", "", "", "minus", "koma", False)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad style"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", -1, "half-up", "minus", False, 1)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad decimals -1"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", 21, "half-up", "minus", False, 1)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad decimals 21"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", "bad", "half-up", "minus", False, 1)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad decimals str"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", 1.5, "half-up", "minus", False, 1)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad decimals 1.5"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", 2, "half-up", "bad", False, 1)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad negativeFormat"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", 2, "bad", "minus", False, 1)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad rounding"',
        "Err.Clear",
        'res = FormatRupiahCore(1000, "Rp", True, ".", ",", 2, "half-up", "minus", True, 7)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:bad compactDecimals 7"',
        "Err.Clear",
        'res = ParseRupiahCore(123)',
        'CheckErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "Err:ParseRupiah non-string"',
        "Err.Clear",
        'res = TerbilangRupiahCore(String(36, "9") & ".999", "lower", "", "", "minus", "rupiah", True, "sen", "half-up")',
        'CheckErr TERBILANG_ERR_OUT_OF_RANGE, Err.Number, "Err:rupiah carry 37 digits"',
        "On Error GoTo 0",
        'WScript.Echo "VBScript Parity & Edge Tests: " & passed & " passed, " & failed & " failed."',
        "If failed > 0 Then WScript.Quit 1",
    ])

    vbs_script = "\r\n".join(test_lines) + "\r\n"

    with tempfile.NamedTemporaryFile("w", suffix=".vbs", encoding="ascii", newline="", delete=False) as tf:
        tf.write(vbs_script)
        tmp_path = Path(tf.name)

    try:
        proc = subprocess.run(
            ["cscript", "//nologo", str(tmp_path)],
            capture_output=True,
            text=True,
            check=False,
        )
        print(proc.stdout.strip())
        if proc.stderr.strip():
            print(proc.stderr.strip(), file=sys.stderr)
        if proc.returncode != 0:
            raise SystemExit(proc.returncode)
        if "VBScript Parity & Edge Tests:" not in proc.stdout or ", 0 failed." not in proc.stdout:
            raise SystemExit(1)
    finally:
        tmp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
