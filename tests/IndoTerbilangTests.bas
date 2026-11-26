Attribute VB_Name = "IndoTerbilangTests"
Option Explicit

' ============================================================================
' Unit Test Suite Native VBA untuk IndoTerbilang
' Impor bersama src/IndoTerbilang.bas lalu jalankan RunAllIndoTerbilangTests.
' ============================================================================

Private mPassed As Long
Private mFailed As Long

Public Sub RunAllIndoTerbilangTests()
    mPassed = 0
    mFailed = 0

    TestTerbilangCore
    TestTerbilangDecimalsAndStyle
    TestTerbilangRupiah
    TestFormatRupiah
    TestParseRupiah
    TestErrors

    Debug.Print "============================================"
    Debug.Print "IndoTerbilang VBA Tests: " & CStr(mPassed) & " passed, " & CStr(mFailed) & " failed."
    Debug.Print "============================================"

    If mFailed > 0 Then
        Err.Raise vbObjectError + 1999, "IndoTerbilangTests", _
            CStr(mFailed) & " test gagal! Lihat Immediate Window (Ctrl+G)."
    Else
        MsgBox "Semua " & CStr(mPassed) & " assertion VBA lulus!", vbInformation, "IndoTerbilangTests"
    End If
End Sub

Private Sub AssertEqualStr(ByVal expected As String, ByVal actual As Variant, ByVal label As String)
    If IsObject(actual) Then
        mFailed = mFailed + 1
        Debug.Print "FAIL [" & label & "] Expected: """ & expected & """ Actual: <Object>"
    ElseIf IsError(actual) Then
        mFailed = mFailed + 1
        Debug.Print "FAIL [" & label & "] Expected: """ & expected & """ Actual: <Error " & CStr(CLng(actual)) & ">"
    ElseIf VarType(actual) = vbString And CStr(actual) = expected Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        Debug.Print "FAIL [" & label & "] Expected: """ & expected & """ Actual: """ & CStr(actual) & """"
    End If
End Sub

Private Sub AssertRaisesErr(ByVal expectedErrNum As Long, ByVal actualErrNum As Long, ByVal label As String)
    If actualErrNum = expectedErrNum Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        Debug.Print "FAIL [" & label & "] Expected Err: " & CStr(expectedErrNum) & " Actual Err: " & CStr(actualErrNum)
    End If
End Sub

Private Sub TestTerbilangCore()
    AssertEqualStr "nol", Terbilang(0), "zero"
    AssertEqualStr "nol", Terbilang(-0#), "negative zero"
    AssertEqualStr "nol", Terbilang("-0.00", , , , , , True), "negative zero string trimmed"
    AssertEqualStr "satu", Terbilang(1), "1"
    AssertEqualStr "sepuluh", Terbilang(10), "10"
    AssertEqualStr "sebelas", Terbilang(11), "11"
    AssertEqualStr "dua belas", Terbilang(12), "12"
    AssertEqualStr "sembilan belas", Terbilang(19), "19"
    AssertEqualStr "dua puluh satu", Terbilang(21), "21"
    AssertEqualStr "seratus", Terbilang(100), "100"
    AssertEqualStr "seratus sebelas", Terbilang(111), "111"
    AssertEqualStr "seribu", Terbilang(1000), "1000"
    AssertEqualStr "seribu satu", Terbilang(1001), "1001"
    AssertEqualStr "dua puluh satu ribu", Terbilang(21000), "21000"
    AssertEqualStr "seratus satu ribu", Terbilang(101000), "101000"
    AssertEqualStr "satu juta", Terbilang(1000000), "1000000"
    AssertEqualStr "satu juta seribu", Terbilang(1001000), "1001000"
    AssertEqualStr "satu miliar", Terbilang(1000000000#), "1e9"
    AssertEqualStr "satu triliun", Terbilang(1000000000000#), "1e12"
    AssertEqualStr "satu kuadriliun", Terbilang("1" & String$(15, "0")), "1e15"
    AssertEqualStr "satu desiliun", Terbilang("1" & String$(33, "0")), "1e33"
    AssertEqualStr "seratus desiliun", Terbilang("1" & String$(35, "0")), "1e35"
    AssertEqualStr "satu sekstiliun", Terbilang(1E+21), "float 1e21"
    AssertEqualStr "lima ratus", Terbilang(".5e3"), "sci .5e3"
    AssertEqualStr "minus seribu lima ratus", Terbilang(-1500), "negative"
    AssertEqualStr "negatif seribu lima ratus", Terbilang(-1500, , , , "negatif"), "custom negativePrefix"
    AssertEqualStr "seribu dua ratus tiga puluh empat", Terbilang(CCur("1234")), "currency integer"
    AssertEqualStr "seratus dua puluh tiga", Terbilang("  123" & vbLf), "trimmed outer newline"
End Sub

Private Sub TestTerbilangDecimalsAndStyle()
    AssertEqualStr "tiga koma satu empat", Terbilang(3.14), "3.14"
    AssertEqualStr "dua belas koma nol lima", Terbilang("12.05"), "12.05"
    AssertEqualStr "dua belas koma lima nol", Terbilang("12.50"), "12.50"
    AssertEqualStr "nol koma lima", Terbilang(".5"), ".5"
    AssertEqualStr "nol koma lima", Terbilang(0.5), "0.5 double (Str$ omits 0)"
    AssertEqualStr "minus nol koma lima", Terbilang(-0.5), "-0.5 double"
    AssertEqualStr "dua belas koma lima", Terbilang("12.50", , , , , , True), "trimTrailingDecimalZeros"
    AssertEqualStr "dua belas", Terbilang("12.00", , , , , , True), "trimTrailingDecimalZeros all zero"
    AssertEqualStr "tiga titik satu empat", Terbilang("3.14", , , , , "titik"), "decimalWord"
    AssertEqualStr "Terbilang: Satu Juta Lima Ratus Ribu Rupiah", _
        Terbilang(1500000, "title", "Terbilang:", "rupiah"), "style title + affixes"
    AssertEqualStr "DUA PULUH LIMA", Terbilang(25, "upper"), "style upper"
    AssertEqualStr "Dua puluh lima", Terbilang(25, "sentence"), "style sentence"
End Sub

Private Sub TestTerbilangRupiah()
    AssertEqualStr "nol rupiah", TerbilangRupiah(0), "0 rupiah"
    AssertEqualStr "seribu rupiah", TerbilangRupiah(1000), "1000 rupiah"
    AssertEqualStr "seribu dua ratus lima puluh rupiah lima puluh sen", _
        TerbilangRupiah(1250.5), "1250.5"
    AssertEqualStr "seribu dua ratus lima puluh rupiah lima sen", _
        TerbilangRupiah("1250.05"), "1250.05"
    AssertEqualStr "seribu dua ratus lima puluh rupiah", _
        TerbilangRupiah("1250.00"), "1250.00"
    AssertEqualStr "nol rupiah lima puluh sen", TerbilangRupiah("0.50"), "0.50"
    AssertEqualStr "seribu dua ratus lima puluh rupiah lima puluh tujuh sen", _
        TerbilangRupiah("1250.567"), "half-up sen"
    AssertEqualStr "seribu dua ratus lima puluh satu rupiah", _
        TerbilangRupiah("1250.996"), "half-up carry"
    AssertEqualStr "seribu dua ratus lima puluh rupiah sembilan puluh sembilan sen", _
        TerbilangRupiah("1250.996", , , , , , , , "truncate"), "truncate sen"
    AssertEqualStr "nol rupiah", TerbilangRupiah("-0.001"), "-0.001 rounds to zero"
    AssertEqualStr "seribu dua ratus lima puluh satu rupiah", _
        TerbilangRupiah(1250.6, , , , , , False), "includeSen false"
    AssertEqualStr "Jumlah: Dua Ribu Lima Ratus Idr Dua Puluh Lima Sen Sah", _
        TerbilangRupiah("2500.25", "title", "Jumlah:", "sah", , "IDR", True, "Sen"), "custom units"
End Sub

Private Sub TestFormatRupiah()
    AssertEqualStr "Rp 0,00", FormatRupiah(0), "fmt 0"
    AssertEqualStr "Rp 0,00", FormatRupiah(-0#), "fmt -0"
    AssertEqualStr "Rp 1.500.000,00", FormatRupiah(1500000), "fmt 1.5M"
    AssertEqualStr "Rp 1.250,57", FormatRupiah("1250.567"), "fmt rounding"
    AssertEqualStr "Rp 1.500.000", FormatRupiah(1500000, , , , , 0), "fmt 0 dec"
    AssertEqualStr "Rp 1.500,5", FormatRupiah("1500.50", , , , , "auto"), "fmt auto"
    AssertEqualStr "IDR1,500.25", FormatRupiah(1500.25, "IDR", False, ",", "."), "fmt custom sep"
    AssertEqualStr "-Rp 1.500.000,00", FormatRupiah(-1500000), "fmt minus"
    AssertEqualStr "(Rp 1.500.000,00)", FormatRupiah(-1500000, , , , , , , "parentheses"), "fmt parentheses"
    AssertEqualStr "Rp 0,00", FormatRupiah("-0.001", , , , , , , "parentheses"), "fmt -0.001"
    AssertEqualStr "Rp 500", FormatRupiah(500, , , , , , , , True), "compact 500"
    AssertEqualStr "Rp 1,5 rb", FormatRupiah(1500, , , , , , , , True), "compact 1.5 rb"
    AssertEqualStr "Rp 1,5 jt", FormatRupiah(1500000, , , , , , , , True), "compact 1.5 jt"
    AssertEqualStr "Rp 2,35 M", FormatRupiah(2350000000#, , , , , , , , True, 2), "compact 2.35 M"
    AssertEqualStr "Rp 1 rb", FormatRupiah("999.96", , , , , , , , True, 1), "compact promote rb"
    AssertEqualStr "Rp 1 jt", FormatRupiah(999990, , , , , , , , True, 1), "compact promote jt"
    AssertEqualStr "Rp 1.500 T", FormatRupiah("1500000000000000", , , , , , , , True), "compact 1500 T"
End Sub

Private Sub TestParseRupiah()
    AssertEqualStr "1500000.50", ParseRupiah("Rp 1.500.000,50"), "parse standard"
    AssertEqualStr "1500000", ParseRupiah("Rp1.500.000"), "parse no space"
    AssertEqualStr "2500000.75", ParseRupiah("IDR 2.500.000,75"), "parse IDR"
    AssertEqualStr "-500000.25", ParseRupiah("(Rp 500.000,25)"), "parse parentheses"
    AssertEqualStr "0.00", ParseRupiah("(Rp 0,00)"), "parse zero parentheses"
    AssertEqualStr "satu juta lima ratus ribu rupiah lima puluh sen", _
        TerbilangRupiah(ParseRupiah("Rp 1.500.000,50")), "parse -> terbilangRupiah"
End Sub

Private Sub TestErrors()
    Dim v As Variant

    On Error Resume Next
    v = Terbilang("")
    AssertRaisesErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "empty string"
    Err.Clear

    v = Terbilang(True)
    AssertRaisesErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "boolean input"
    Err.Clear

    v = Terbilang("1.500.000")
    AssertRaisesErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "grouped dots"
    Err.Clear

    v = Terbilang("1" & String$(36, "0"))
    AssertRaisesErr TERBILANG_ERR_OUT_OF_RANGE, Err.Number, "37 digits"
    Err.Clear

    v = Terbilang("1e101")
    AssertRaisesErr TERBILANG_ERR_OUT_OF_RANGE, Err.Number, "exp 101"
    Err.Clear

    v = TerbilangRupiah(String$(36, "9") & ".999")
    AssertRaisesErr TERBILANG_ERR_OUT_OF_RANGE, Err.Number, "rounding carry 37 digits"
    Err.Clear

    v = FormatRupiah(1000, , , , , -1)
    AssertRaisesErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "bad decimals"
    Err.Clear

    v = ParseRupiah("Rp 1.50.000")
    AssertRaisesErr TERBILANG_ERR_INVALID_INPUT, Err.Number, "bad rp grouping"
    Err.Clear
    On Error GoTo 0
End Sub
