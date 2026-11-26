Attribute VB_Name = "IndoTerbilang"
Option Explicit

' ============================================================================
' indo-terbilang (Excel VBA / Macro) v1.0.0
' Zero-dependency Indonesian number-to-words (EYD V / KBBI) and
' deterministic Rupiah currency formatter & parser for Microsoft Excel.
'
' License : MIT
' Author  : salmankalista
' ============================================================================

' Public structured error numbers when called from VBA macros (Sub/Function)
Public Const TERBILANG_ERR_INVALID_INPUT As Long = vbObjectError + 1001
Public Const TERBILANG_ERR_NON_FINITE As Long = vbObjectError + 1002
Public Const TERBILANG_ERR_OUT_OF_RANGE As Long = vbObjectError + 1003

Private Const MAX_INTEGER_DIGITS As Long = 36
Private Const MAX_EXPONENT_ABS As Long = 100
Private Const VT_LONGLONG As Long = 20
Private Const XL_ERR_VALUE As Long = 2015
Private Const XL_ERR_NUM As Long = 2036

' ============================================================================
' PUBLIC API (Worksheet Formulas / UDF & VBA Functions)
' ============================================================================

Public Function Terbilang( _
    ByVal Value As Variant, _
    Optional ByVal Style As String = "lower", _
    Optional ByVal Prefix As String = "", _
    Optional ByVal Suffix As String = "", _
    Optional ByVal NegativePrefix As String = "minus", _
    Optional ByVal DecimalWord As String = "koma", _
    Optional ByVal TrimTrailingDecimalZeros As Boolean = False _
) As Variant
    Dim unwrapped As Variant
    On Error GoTo HandleErr

    UnwrapInputValue Value, unwrapped
    If IsError(unwrapped) Then
        If IsWorksheetCall() Then
            Terbilang = unwrapped
            Exit Function
        End If
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input berisi nilai error Excel."
    End If

    Terbilang = TerbilangCore( _
        unwrapped, Style, Prefix, Suffix, NegativePrefix, DecimalWord, TrimTrailingDecimalZeros _
    )
    Exit Function

HandleErr:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    If IsWorksheetCall() Then
        Terbilang = MapErrToExcelError(errNum)
    Else
        Err.Raise errNum, errSrc, errDesc
    End If
End Function

Public Function TerbilangRupiah( _
    ByVal Value As Variant, _
    Optional ByVal Style As String = "lower", _
    Optional ByVal Prefix As String = "", _
    Optional ByVal Suffix As String = "", _
    Optional ByVal NegativePrefix As String = "minus", _
    Optional ByVal CurrencyUnit As String = "rupiah", _
    Optional ByVal IncludeSen As Boolean = True, _
    Optional ByVal SenUnit As String = "sen", _
    Optional ByVal Rounding As String = "half-up" _
) As Variant
    Dim unwrapped As Variant
    On Error GoTo HandleErr

    UnwrapInputValue Value, unwrapped
    If IsError(unwrapped) Then
        If IsWorksheetCall() Then
            TerbilangRupiah = unwrapped
            Exit Function
        End If
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input berisi nilai error Excel."
    End If

    TerbilangRupiah = TerbilangRupiahCore( _
        unwrapped, Style, Prefix, Suffix, NegativePrefix, CurrencyUnit, IncludeSen, SenUnit, Rounding _
    )
    Exit Function

HandleErr:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    If IsWorksheetCall() Then
        TerbilangRupiah = MapErrToExcelError(errNum)
    Else
        Err.Raise errNum, errSrc, errDesc
    End If
End Function

Public Function FormatRupiah( _
    ByVal Value As Variant, _
    Optional ByVal Symbol As String = "Rp", _
    Optional ByVal SpaceAfterSymbol As Boolean = True, _
    Optional ByVal ThousandsSeparator As String = ".", _
    Optional ByVal DecimalSeparator As String = ",", _
    Optional ByVal Decimals As Variant = 2, _
    Optional ByVal Rounding As String = "half-up", _
    Optional ByVal NegativeFormat As String = "minus", _
    Optional ByVal Compact As Boolean = False, _
    Optional ByVal CompactDecimals As Variant = 1 _
) As Variant
    Dim unwrapped As Variant
    On Error GoTo HandleErr

    UnwrapInputValue Value, unwrapped
    If IsError(unwrapped) Then
        If IsWorksheetCall() Then
            FormatRupiah = unwrapped
            Exit Function
        End If
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input berisi nilai error Excel."
    End If

    FormatRupiah = FormatRupiahCore( _
        unwrapped, Symbol, SpaceAfterSymbol, ThousandsSeparator, DecimalSeparator, _
        Decimals, Rounding, NegativeFormat, Compact, CompactDecimals _
    )
    Exit Function

HandleErr:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    If IsWorksheetCall() Then
        FormatRupiah = MapErrToExcelError(errNum)
    Else
        Err.Raise errNum, errSrc, errDesc
    End If
End Function

Public Function ParseRupiah(ByVal Value As Variant) As Variant
    Dim unwrapped As Variant
    On Error GoTo HandleErr

    UnwrapInputValue Value, unwrapped
    If IsError(unwrapped) Then
        If IsWorksheetCall() Then
            ParseRupiah = unwrapped
            Exit Function
        End If
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input berisi nilai error Excel."
    End If

    ParseRupiah = ParseRupiahCore(unwrapped)
    Exit Function

HandleErr:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    If IsWorksheetCall() Then
        ParseRupiah = MapErrToExcelError(errNum)
    Else
        Err.Raise errNum, errSrc, errDesc
    End If
End Function

' Optional utility: register function descriptions in Excel Insert Function (fx) dialog
Public Sub RegisterIndoTerbilangFunctions()
    Dim argsTerbilang As Variant
    Dim argsRupiah As Variant
    Dim argsFormat As Variant
    Dim argsParse As Variant

    argsTerbilang = Array( _
        "Angka, sel, atau teks numerik (maks. 36 digit). Contoh: 1500000 atau ""12.50"".", _
        "[Opsional] Gaya huruf: ""lower"" (default), ""upper"", ""title"", atau ""sentence"".", _
        "[Opsional] Teks awalan sebelum kalimat hasil. Contoh: ""Terbilang:"".", _
        "[Opsional] Teks akhiran setelah kalimat hasil. Contoh: ""rupiah"".", _
        "[Opsional] Kata untuk tanda minus. Default: ""minus"" (mis. ""negatif"").", _
        "[Opsional] Kata pemisah desimal. Default: ""koma"" (mis. ""titik"").", _
        "[Opsional] TRUE untuk buang nol di akhir desimal teks. Default: FALSE." _
    )
    TryRegisterMacroOption "Terbilang", _
        "Mengonversi angka menjadi kalimat terbilang Bahasa Indonesia baku (EYD V / KBBI).", _
        "IndoTerbilang", argsTerbilang

    argsRupiah = Array( _
        "Nominal angka, sel, atau teks numerik. Contoh: 1250750.5 atau ""1250.50"".", _
        "[Opsional] Gaya huruf: ""lower"" (default), ""upper"", ""title"", atau ""sentence"".", _
        "[Opsional] Teks awalan sebelum kalimat hasil. Contoh: ""Jumlah:"".", _
        "[Opsional] Teks akhiran setelah satuan mata uang/sen. Contoh: ""sah"".", _
        "[Opsional] Kata untuk tanda minus. Default: ""minus"".", _
        "[Opsional] Satuan mata uang utama. Default: ""rupiah"".", _
        "[Opsional] TRUE (default) untuk menyertakan sen jika > 0; FALSE untuk dibulatkan.", _
        "[Opsional] Satuan pecahan desimal. Default: ""sen"".", _
        "[Opsional] Pembulatan >2 desimal: ""half-up"" (default) atau ""truncate""." _
    )
    TryRegisterMacroOption "TerbilangRupiah", _
        "Mengonversi nominal angka menjadi kalimat terbilang Rupiah beserta sen.", _
        "IndoTerbilang", argsRupiah

    argsFormat = Array( _
        "Nominal angka, sel, atau teks numerik. Contoh: 1500000 atau ""1250.567"".", _
        "[Opsional] Simbol mata uang. Default: ""Rp"" (isi """" jika tanpa simbol).", _
        "[Opsional] TRUE (default) untuk spasi setelah simbol (""Rp 1.000,00"").", _
        "[Opsional] Pemisah ribuan. Default: ""."".", _
        "[Opsional] Pemisah desimal. Default: "","".", _
        "[Opsional] Jumlah desimal 0-20 (default: 2) atau ""auto"".", _
        "[Opsional] Mode pembulatan: ""half-up"" (default) atau ""truncate"".", _
        "[Opsional] Format negatif: ""minus"" (default) atau ""parentheses"".", _
        "[Opsional] TRUE untuk notasi ringkas (rb, jt, M, T). Default: FALSE.", _
        "[Opsional] Maks. digit desimal saat mode compact (0-6). Default: 1." _
    )
    TryRegisterMacroOption "FormatRupiah", _
        "Memformat nominal menjadi string Rupiah secara deterministik atau ringkas (compact).", _
        "IndoTerbilang", argsFormat

    argsParse = Array( _
        "Teks berformat Rupiah. Contoh: ""Rp 1.500.000,50"" atau ""(Rp 500.000)""." _
    )
    TryRegisterMacroOption "ParseRupiah", _
        "Mengekstrak teks Rupiah menjadi string numerik kanonik siap hitung.", _
        "IndoTerbilang", argsParse
End Sub

Private Sub TryRegisterMacroOption( _
    ByVal macroName As String, _
    ByVal descriptionText As String, _
    ByVal categoryName As String, _
    ByVal argDescriptions As Variant _
)
    On Error Resume Next
    Application.MacroOptions _
        Macro:=macroName, _
        Description:=descriptionText, _
        Category:=categoryName, _
        ArgumentDescriptions:=argDescriptions
    If Err.Number <> 0 Then
        Err.Clear
        Application.MacroOptions _
            Macro:=macroName, _
            Description:=descriptionText, _
            Category:=categoryName
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' CORE FUNCTIONS
' ============================================================================

Private Function TerbilangCore( _
    ByVal Value As Variant, _
    ByVal Style As String, _
    ByVal Prefix As String, _
    ByVal Suffix As String, _
    ByVal NegativePrefix As String, _
    ByVal DecimalWord As String, _
    ByVal TrimTrailingDecimalZeros As Boolean _
) As String
    Dim isNegative As Boolean
    Dim intStr As String
    Dim fracStr As String
    Dim intWords As String
    Dim fracWords As String
    Dim textOut As String
    Dim i As Long
    Dim digitVal As Long

    ParseNumericInput Value, isNegative, intStr, fracStr

    If TrimTrailingDecimalZeros Then
        fracStr = TrimTrailingZeros(fracStr)
    End If

    intWords = ConvertIntegerToWords(intStr)
    textOut = intWords

    If Len(fracStr) > 0 Then
        fracWords = ""
        For i = 1 To Len(fracStr)
            digitVal = AscW(Mid$(fracStr, i, 1)) - 48
            If i = 1 Then
                fracWords = UnitWord(digitVal)
            Else
                fracWords = fracWords & " " & UnitWord(digitVal)
            End If
        Next i
        textOut = intWords & " " & TrimWhitespace(DecimalWord) & " " & fracWords
    End If

    If isNegative Then
        textOut = TrimWhitespace(NegativePrefix) & " " & textOut
    End If

    TerbilangCore = ApplyStyleAndAffixes(textOut, Style, Prefix, Suffix)
End Function

Private Function TerbilangRupiahCore( _
    ByVal Value As Variant, _
    ByVal Style As String, _
    ByVal Prefix As String, _
    ByVal Suffix As String, _
    ByVal NegativePrefix As String, _
    ByVal CurrencyUnit As String, _
    ByVal IncludeSen As Boolean, _
    ByVal SenUnit As String, _
    ByVal Rounding As String _
) As String
    Dim isNegative As Boolean
    Dim intStr As String
    Dim fracStr As String
    Dim roundedInt As String
    Dim roundedFrac As String
    Dim decimalsCount As Long
    Dim senVal As Long
    Dim isZero As Boolean
    Dim intWords As String
    Dim textOut As String

    ParseNumericInput Value, isNegative, intStr, fracStr

    If IncludeSen Then
        decimalsCount = 2
    Else
        decimalsCount = 0
    End If

    RoundDecimal intStr, fracStr, decimalsCount, Rounding, roundedInt, roundedFrac
    AssertMaxDigits roundedInt

    If IncludeSen And Len(roundedFrac) > 0 Then
        senVal = CLng(roundedFrac)
    Else
        senVal = 0
    End If

    isZero = (roundedInt = "0" And senVal = 0)
    If isZero Then
        isNegative = False
    End If

    intWords = ConvertIntegerToWords(roundedInt)
    textOut = intWords & " " & TrimWhitespace(CurrencyUnit)

    If senVal > 0 Then
        textOut = textOut & " " & ConvertTriplet(senVal) & " " & TrimWhitespace(SenUnit)
    End If

    If isNegative Then
        textOut = TrimWhitespace(NegativePrefix) & " " & textOut
    End If

    TerbilangRupiahCore = ApplyStyleAndAffixes(textOut, Style, Prefix, Suffix)
End Function

Private Function FormatRupiahCore( _
    ByVal Value As Variant, _
    ByVal Symbol As String, _
    ByVal SpaceAfterSymbol As Boolean, _
    ByVal ThousandsSeparator As String, _
    ByVal DecimalSeparator As String, _
    ByVal Decimals As Variant, _
    ByVal Rounding As String, _
    ByVal NegativeFormat As String, _
    ByVal Compact As Boolean, _
    ByVal CompactDecimals As Variant _
) As String
    Dim normNegFormat As String
    Dim normRounding As String
    Dim isNegative As Boolean
    Dim intStr As String
    Dim fracStr As String
    Dim compactDecCount As Long
    Dim decCount As Long
    Dim isAutoDecimals As Boolean
    Dim finalInt As String
    Dim finalFrac As String
    Dim finalIsNeg As Boolean
    Dim isZero As Boolean
    Dim formattedInt As String
    Dim numberText As String

    normNegFormat = LCase$(TrimWhitespace(NegativeFormat))
    If normNegFormat <> "minus" And normNegFormat <> "parentheses" Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
            "Opsi NegativeFormat tidak valid: """ & NegativeFormat & """."
    End If

    normRounding = NormalizeRoundingMode(Rounding)
    ParseNumericInput Value, isNegative, intStr, fracStr

    If Compact Then
        ParseWholeNumberOption CompactDecimals, 0, 6, False, "CompactDecimals", compactDecCount, isAutoDecimals
        FormatRupiahCore = FormatCompactRupiah( _
            isNegative, intStr, fracStr, Symbol, SpaceAfterSymbol, _
            ThousandsSeparator, DecimalSeparator, normRounding, normNegFormat, compactDecCount _
        )
        Exit Function
    End If

    ParseWholeNumberOption Decimals, 0, 20, True, "Decimals", decCount, isAutoDecimals

    If isAutoDecimals Then
        finalInt = intStr
        finalFrac = TrimTrailingZeros(fracStr)
        finalIsNeg = isNegative
    Else
        RoundDecimal intStr, fracStr, decCount, normRounding, finalInt, finalFrac
        isZero = (finalInt = "0" And IsZeroOnly(finalFrac))
        If isZero Then
            finalIsNeg = False
        Else
            finalIsNeg = isNegative
        End If
    End If

    formattedInt = InsertThousandsSeparator(finalInt, ThousandsSeparator)
    If Len(finalFrac) > 0 Then
        numberText = formattedInt & DecimalSeparator & finalFrac
    Else
        numberText = formattedInt
    End If

    FormatRupiahCore = ApplyRupiahAffixes( _
        numberText, finalIsNeg, Symbol, SpaceAfterSymbol, normNegFormat _
    )
End Function

Private Function FormatCompactRupiah( _
    ByVal isNegativeInput As Boolean, _
    ByVal integerStr As String, _
    ByVal fractionStr As String, _
    ByVal Symbol As String, _
    ByVal SpaceAfterSymbol As Boolean, _
    ByVal ThousandsSeparator As String, _
    ByVal DecimalSeparator As String, _
    ByVal normRounding As String, _
    ByVal normNegFormat As String, _
    ByVal compactDecimals As Long _
) As String
    Dim powers(0 To 3) As Long
    Dim labels(0 To 3) As String
    Dim scaleIdx As Long
    Dim i As Long
    Dim roundedInt As String
    Dim roundedFrac As String
    Dim unitLabel As String
    Dim finalFrac As String
    Dim isZero As Boolean
    Dim isNeg As Boolean
    Dim formattedInt As String
    Dim numberPart As String

    powers(0) = 12: labels(0) = "T"
    powers(1) = 9:  labels(1) = "M"
    powers(2) = 6:  labels(2) = "jt"
    powers(3) = 3:  labels(3) = "rb"

    scaleIdx = -1
    For i = 0 To 3
        If Len(integerStr) > powers(i) Then
            scaleIdx = i
            Exit For
        End If
    Next i

    unitLabel = ""
    If scaleIdx = -1 Then
        RoundDecimal integerStr, fractionStr, compactDecimals, normRounding, roundedInt, roundedFrac
        If Len(roundedInt) > 3 Then
            scaleIdx = 3
        End If
    End If

    If scaleIdx <> -1 Then
        DivideByPowerAndRound _
            integerStr, fractionStr, powers(scaleIdx), compactDecimals, normRounding, roundedInt, roundedFrac
        If Len(roundedInt) > 3 And scaleIdx > 0 Then
            scaleIdx = scaleIdx - 1
            DivideByPowerAndRound _
                integerStr, fractionStr, powers(scaleIdx), compactDecimals, normRounding, roundedInt, roundedFrac
        End If
        unitLabel = " " & labels(scaleIdx)
    End If

    finalFrac = TrimTrailingZeros(roundedFrac)
    isZero = (roundedInt = "0" And Len(finalFrac) = 0)
    If isZero Then
        isNeg = False
    Else
        isNeg = isNegativeInput
    End If

    formattedInt = InsertThousandsSeparator(roundedInt, ThousandsSeparator)
    If Len(finalFrac) > 0 Then
        numberPart = formattedInt & DecimalSeparator & finalFrac
    Else
        numberPart = formattedInt
    End If

    FormatCompactRupiah = ApplyRupiahAffixes( _
        numberPart & unitLabel, isNeg, Symbol, SpaceAfterSymbol, normNegFormat _
    )
End Function

Private Function ParseRupiahCore(ByVal Value As Variant) As String
    Dim unwrapped As Variant
    Dim raw As String
    Dim isNegative As Boolean
    Dim commaIdx As Long
    Dim intRaw As String
    Dim fracRaw As String
    Dim cleanInt As String
    Dim normalizedInt As String
    Dim isZero As Boolean
    Dim signPrefix As String

    UnwrapInputValue Value, unwrapped
    If VarType(unwrapped) <> vbString Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input harus string."
    End If

    raw = TrimWhitespace(CStr(unwrapped))
    If Len(raw) = 0 Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input kosong."
    End If

    isNegative = False
    If Left$(raw, 1) = "(" And Right$(raw, 1) = ")" Then
        isNegative = True
        raw = TrimWhitespace(Mid$(raw, 2, Len(raw) - 2))
        If Left$(raw, 1) = "-" Or Left$(raw, 1) = "+" Then
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Tanda negatif ganda tidak valid."
        End If
    ElseIf Left$(raw, 1) = "-" Then
        isNegative = True
        raw = TrimWhitespace(Mid$(raw, 2))
    ElseIf Left$(raw, 1) = "+" Then
        raw = TrimWhitespace(Mid$(raw, 2))
    End If

    If LCase$(Left$(raw, 3)) = "idr" Then
        raw = TrimWhitespace(Mid$(raw, 4))
    ElseIf LCase$(Left$(raw, 2)) = "rp" Then
        raw = TrimWhitespace(Mid$(raw, 3))
    End If

    commaIdx = InStr(1, raw, ",", vbBinaryCompare)
    If commaIdx > 0 Then
        intRaw = Left$(raw, commaIdx - 1)
        fracRaw = Mid$(raw, commaIdx + 1)
        If Not IsAllAsciiDigits(fracRaw) Then
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Bagian desimal Rupiah tidak valid."
        End If
    Else
        intRaw = raw
        fracRaw = ""
    End If

    If InStr(1, intRaw, ".", vbBinaryCompare) > 0 Then
        If Not IsValidGroupedThousands(intRaw) Then
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Pemisah ribuan tidak valid."
        End If
        cleanInt = Replace(intRaw, ".", "")
    Else
        If Not IsAllAsciiDigits(intRaw) Then
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Format Rupiah tidak valid."
        End If
        cleanInt = intRaw
    End If

    normalizedInt = TrimLeadingZeros(cleanInt)
    AssertMaxDigits normalizedInt

    isZero = (normalizedInt = "0" And IsZeroOnly(fracRaw))
    If isNegative And Not isZero Then
        signPrefix = "-"
    Else
        signPrefix = ""
    End If

    If Len(fracRaw) > 0 Then
        ParseRupiahCore = signPrefix & normalizedInt & "." & fracRaw
    Else
        ParseRupiahCore = signPrefix & normalizedInt
    End If
End Function

' ============================================================================
' WORD CONVERSION (EYD V / KBBI)
' ============================================================================

Private Function ConvertTriplet(ByVal n As Long) As String
    Dim words As String
    Dim hundreds As Long
    Dim remainder As Long
    Dim tens As Long
    Dim ones As Long

    words = ""
    hundreds = n \ 100
    remainder = n Mod 100

    If hundreds = 1 Then
        AppendWord words, "seratus"
    ElseIf hundreds > 1 Then
        AppendWord words, UnitWord(hundreds)
        AppendWord words, "ratus"
    End If

    If remainder > 0 Then
        If remainder < 10 Then
            AppendWord words, UnitWord(remainder)
        ElseIf remainder = 10 Then
            AppendWord words, "sepuluh"
        ElseIf remainder = 11 Then
            AppendWord words, "sebelas"
        ElseIf remainder < 20 Then
            AppendWord words, UnitWord(remainder - 10)
            AppendWord words, "belas"
        Else
            tens = remainder \ 10
            ones = remainder Mod 10
            AppendWord words, UnitWord(tens)
            AppendWord words, "puluh"
            If ones > 0 Then
                AppendWord words, UnitWord(ones)
            End If
        End If
    End If

    ConvertTriplet = words
End Function

Private Function ConvertIntegerToWords(ByVal integerStr As String) As String
    Dim endPos As Long
    Dim startPos As Long
    Dim chunkVal As Long
    Dim scaleIdx As Long
    Dim seg As String
    Dim scaleWord As String
    Dim resultStr As String

    If integerStr = "0" Then
        ConvertIntegerToWords = "nol"
        Exit Function
    End If

    resultStr = ""
    scaleIdx = 0
    endPos = Len(integerStr)

    Do While endPos > 0
        startPos = endPos - 2
        If startPos < 1 Then startPos = 1
        chunkVal = CLng(Mid$(integerStr, startPos, endPos - startPos + 1))

        If chunkVal > 0 Then
            If scaleIdx = 1 And chunkVal = 1 Then
                seg = "seribu"
            Else
                seg = ConvertTriplet(chunkVal)
                scaleWord = ScaleWordAt(scaleIdx)
                If Len(scaleWord) > 0 Then
                    seg = seg & " " & scaleWord
                End If
            End If

            If Len(resultStr) = 0 Then
                resultStr = seg
            Else
                resultStr = seg & " " & resultStr
            End If
        End If

        scaleIdx = scaleIdx + 1
        endPos = startPos - 1
    Loop

    ConvertIntegerToWords = resultStr
End Function

Private Function UnitWord(ByVal digit As Long) As String
    Select Case digit
        Case 0: UnitWord = "nol"
        Case 1: UnitWord = "satu"
        Case 2: UnitWord = "dua"
        Case 3: UnitWord = "tiga"
        Case 4: UnitWord = "empat"
        Case 5: UnitWord = "lima"
        Case 6: UnitWord = "enam"
        Case 7: UnitWord = "tujuh"
        Case 8: UnitWord = "delapan"
        Case 9: UnitWord = "sembilan"
    End Select
End Function

Private Function ScaleWordAt(ByVal idx As Long) As String
    Select Case idx
        Case 0:  ScaleWordAt = ""
        Case 1:  ScaleWordAt = "ribu"
        Case 2:  ScaleWordAt = "juta"
        Case 3:  ScaleWordAt = "miliar"
        Case 4:  ScaleWordAt = "triliun"
        Case 5:  ScaleWordAt = "kuadriliun"
        Case 6:  ScaleWordAt = "kuintiliun"
        Case 7:  ScaleWordAt = "sekstiliun"
        Case 8:  ScaleWordAt = "septiliun"
        Case 9:  ScaleWordAt = "oktiliun"
        Case 10: ScaleWordAt = "noniliun"
        Case 11: ScaleWordAt = "desiliun"
    End Select
End Function

' ============================================================================
' PARSER, ROUNDING, AND STYLING HELPERS
' ============================================================================

Private Sub UnwrapInputValue(ByVal Value As Variant, ByRef OutValue As Variant)
    If IsObject(Value) Then
        If Value Is Nothing Then
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Objek Nothing tidak didukung."
        End If
        If TypeName(Value) = "Range" Then
            If Value.Areas.Count > 1 Or Value.Rows.Count > 1 Or Value.Columns.Count > 1 Then
                RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Range multi-sel tidak didukung."
            End If
            OutValue = Value.Value
        Else
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Tipe objek tidak didukung."
        End If
    Else
        OutValue = Value
    End If
End Sub

Private Sub ParseNumericInput( _
    ByVal Value As Variant, _
    ByRef OutIsNegative As Boolean, _
    ByRef OutInteger As String, _
    ByRef OutFraction As String _
)
    Dim v As Variant
    Dim vt As Long
    Dim raw As String
    Dim isNegative As Boolean
    Dim body As String
    Dim ePos As Long
    Dim mantissa As String
    Dim expRaw As String
    Dim expVal As Long
    Dim intPart As String
    Dim fracPart As String
    Dim trimmedInt As String

    UnwrapInputValue Value, v
    vt = VarType(v)

    If (vt And vbArray) <> 0 Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input array tidak didukung."
    End If

    Select Case vt
        Case vbByte, vbInteger, vbLong, VT_LONGLONG, vbSingle, vbDouble, vbCurrency, vbDecimal
            If vt = vbSingle Or vt = vbDouble Then
                If v = 0 Then
                    OutIsNegative = False
                    OutInteger = "0"
                    OutFraction = ""
                    Exit Sub
                End If
            End If
            ' Str$ always uses "." as decimal separator regardless of Windows regional settings
            raw = Trim$(Str$(v))
            If InStr(1, raw, "#", vbBinaryCompare) > 0 _
                Or InStr(1, UCase$(raw), "INF", vbBinaryCompare) > 0 _
                Or InStr(1, UCase$(raw), "NAN", vbBinaryCompare) > 0 Then
                RaiseTerbilangError TERBILANG_ERR_NON_FINITE, "NON_FINITE_NUMBER", "Angka harus finite."
            End If

        Case vbString
            raw = TrimWhitespace(CStr(v))
            If Len(raw) = 0 Then
                RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Input kosong."
            End If

        Case Else
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Tipe input tidak didukung."
    End Select

    isNegative = False
    body = raw
    If Left$(body, 1) = "-" Then
        isNegative = True
        body = Mid$(body, 2)
    ElseIf Left$(body, 1) = "+" Then
        body = Mid$(body, 2)
    End If

    If Len(body) = 0 Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Format angka tidak valid."
    End If

    ePos = InStr(1, body, "e", vbTextCompare)
    If ePos > 0 Then
        mantissa = NormalizeDecimalSeparator(Left$(body, ePos - 1))
        expRaw = Mid$(body, ePos + 1)
        expVal = ParseExponent(expRaw)
        ExpandScientific mantissa, expVal, intPart, fracPart
    Else
        SplitStandardDecimal NormalizeDecimalSeparator(body), intPart, fracPart
    End If

    trimmedInt = TrimLeadingZeros(intPart)
    AssertMaxDigits trimmedInt

    If trimmedInt = "0" And IsZeroOnly(fracPart) Then
        OutIsNegative = False
    Else
        OutIsNegative = isNegative
    End If
    OutInteger = trimmedInt
    OutFraction = fracPart
End Sub

Private Function NormalizeDecimalSeparator(ByVal raw As String) As String
    Dim hasDot As Boolean
    Dim hasComma As Boolean
    Dim normalized As String

    hasDot = (InStr(1, raw, ".", vbBinaryCompare) > 0)
    hasComma = (InStr(1, raw, ",", vbBinaryCompare) > 0)

    If hasDot And hasComma Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Pemisah titik dan koma bercampur."
    End If

    If hasComma Then
        normalized = Replace(raw, ",", ".")
    Else
        normalized = raw
    End If

    If Not IsValidUnsignedDecimal(normalized) Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
            "Format angka tidak valid: """ & raw & """."
    End If

    NormalizeDecimalSeparator = normalized
End Function

Private Function IsValidUnsignedDecimal(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim dotCount As Long
    Dim digitCount As Long

    If Len(s) = 0 Then
        IsValidUnsignedDecimal = False
        Exit Function
    End If
    If Right$(s, 1) = "." Then
        IsValidUnsignedDecimal = False
        Exit Function
    End If

    dotCount = 0
    digitCount = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dotCount = dotCount + 1
            If dotCount > 1 Then
                IsValidUnsignedDecimal = False
                Exit Function
            End If
        Else
            code = AscW(ch)
            If code < 48 Or code > 57 Then
                IsValidUnsignedDecimal = False
                Exit Function
            End If
            digitCount = digitCount + 1
        End If
    Next i

    IsValidUnsignedDecimal = (digitCount > 0)
End Function

Private Function ParseExponent(ByVal expRaw As String) As Long
    Dim signMult As Long
    Dim digitsPart As String
    Dim trimmedExp As String

    If Len(expRaw) = 0 Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Eksponen tidak valid."
    End If

    signMult = 1
    If Left$(expRaw, 1) = "-" Then
        signMult = -1
        digitsPart = Mid$(expRaw, 2)
    ElseIf Left$(expRaw, 1) = "+" Then
        digitsPart = Mid$(expRaw, 2)
    Else
        digitsPart = expRaw
    End If

    If Not IsAllAsciiDigits(digitsPart) Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Eksponen tidak valid."
    End If

    trimmedExp = TrimLeadingZeros(digitsPart)
    If Len(trimmedExp) > 3 Then
        RaiseTerbilangError TERBILANG_ERR_OUT_OF_RANGE, "OUT_OF_RANGE", "Eksponen melebihi batas."
    End If

    ParseExponent = signMult * CLng(trimmedExp)
    If Abs(ParseExponent) > MAX_EXPONENT_ABS Then
        RaiseTerbilangError TERBILANG_ERR_OUT_OF_RANGE, "OUT_OF_RANGE", "Eksponen melebihi batas."
    End If
End Function

Private Sub ExpandScientific( _
    ByVal mantissa As String, _
    ByVal expVal As Long, _
    ByRef OutInteger As String, _
    ByRef OutFraction As String _
)
    Dim dotIdx As Long
    Dim intPart As String
    Dim fracPart As String
    Dim digits As String
    Dim newDotPos As Long

    dotIdx = InStr(1, mantissa, ".", vbBinaryCompare)
    If dotIdx = 0 Then
        intPart = mantissa
        fracPart = ""
    Else
        intPart = Left$(mantissa, dotIdx - 1)
        fracPart = Mid$(mantissa, dotIdx + 1)
    End If

    digits = intPart & fracPart
    newDotPos = Len(intPart) + expVal

    If newDotPos <= 0 Then
        OutInteger = "0"
        OutFraction = String$(-newDotPos, "0") & digits
    ElseIf newDotPos >= Len(digits) Then
        OutInteger = digits & String$(newDotPos - Len(digits), "0")
        OutFraction = ""
    Else
        OutInteger = Left$(digits, newDotPos)
        OutFraction = Mid$(digits, newDotPos + 1)
    End If
End Sub

Private Sub SplitStandardDecimal( _
    ByVal normalized As String, _
    ByRef OutInteger As String, _
    ByRef OutFraction As String _
)
    Dim dotIdx As Long
    Dim intPart As String

    dotIdx = InStr(1, normalized, ".", vbBinaryCompare)
    If dotIdx = 0 Then
        OutInteger = normalized
        OutFraction = ""
    Else
        intPart = Left$(normalized, dotIdx - 1)
        If Len(intPart) = 0 Then
            OutInteger = "0"
        Else
            OutInteger = intPart
        End If
        OutFraction = Mid$(normalized, dotIdx + 1)
    End If
End Sub

Private Sub RoundDecimal( _
    ByVal integerStr As String, _
    ByVal fractionStr As String, _
    ByVal decimalsCount As Long, _
    ByVal mode As String, _
    ByRef OutInteger As String, _
    ByRef OutFraction As String _
)
    Dim normMode As String
    Dim fracLen As Long
    Dim keptFraction As String
    Dim nextDigit As Long
    Dim combined As String
    Dim i As Long
    Dim carry As Long
    Dim total As Long
    Dim splitLen As Long

    If decimalsCount < 0 Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", "Desimal pembulatan tidak valid."
    End If

    normMode = NormalizeRoundingMode(mode)
    fracLen = Len(fractionStr)

    If fracLen = decimalsCount Then
        OutInteger = integerStr
        OutFraction = fractionStr
        Exit Sub
    End If

    If fracLen < decimalsCount Then
        OutInteger = integerStr
        OutFraction = fractionStr & String$(decimalsCount - fracLen, "0")
        Exit Sub
    End If

    keptFraction = Left$(fractionStr, decimalsCount)
    nextDigit = AscW(Mid$(fractionStr, decimalsCount + 1, 1)) - 48

    If normMode = "truncate" Or nextDigit < 5 Then
        OutInteger = integerStr
        OutFraction = keptFraction
        Exit Sub
    End If

    combined = integerStr & keptFraction
    carry = 1
    For i = Len(combined) To 1 Step -1
        If carry = 0 Then Exit For
        total = (AscW(Mid$(combined, i, 1)) - 48) + carry
        If total >= 10 Then
            Mid$(combined, i, 1) = ChrW(48 + total - 10)
        Else
            Mid$(combined, i, 1) = ChrW(48 + total)
            carry = 0
        End If
    Next i

    If carry > 0 Then
        combined = "1" & combined
    End If

    If decimalsCount = 0 Then
        OutInteger = combined
        OutFraction = ""
    Else
        splitLen = Len(combined) - decimalsCount
        OutInteger = Left$(combined, splitLen)
        OutFraction = Right$(combined, decimalsCount)
    End If
End Sub

Private Sub DivideByPowerAndRound( _
    ByVal integerStr As String, _
    ByVal fractionStr As String, _
    ByVal powerVal As Long, _
    ByVal decimalsCount As Long, _
    ByVal mode As String, _
    ByRef OutInteger As String, _
    ByRef OutFraction As String _
)
    Dim splitAt As Long
    Dim shiftedInt As String
    Dim shiftedFrac As String

    splitAt = Len(integerStr) - powerVal
    If splitAt > 0 Then
        shiftedInt = Left$(integerStr, splitAt)
        shiftedFrac = Mid$(integerStr, splitAt + 1) & fractionStr
    Else
        shiftedInt = "0"
        shiftedFrac = String$(powerVal - Len(integerStr), "0") & integerStr & fractionStr
    End If

    RoundDecimal shiftedInt, shiftedFrac, decimalsCount, mode, OutInteger, OutFraction
End Sub

Private Function ApplyStyleAndAffixes( _
    ByVal textStr As String, _
    ByVal Style As String, _
    ByVal Prefix As String, _
    ByVal Suffix As String _
) As String
    Dim cleanPrefix As String
    Dim cleanSuffix As String
    Dim combined As String
    Dim normStyle As String
    Dim words() As String
    Dim i As Long
    Dim w As String

    combined = textStr
    cleanPrefix = TrimWhitespace(Prefix)
    If Len(cleanPrefix) > 0 Then
        combined = cleanPrefix & " " & combined
    End If

    cleanSuffix = TrimWhitespace(Suffix)
    If Len(cleanSuffix) > 0 Then
        combined = combined & " " & cleanSuffix
    End If

    normStyle = LCase$(TrimWhitespace(Style))
    Select Case normStyle
        Case "lower"
            ApplyStyleAndAffixes = LCase$(combined)
        Case "upper"
            ApplyStyleAndAffixes = UCase$(combined)
        Case "title"
            words = Split(combined, " ")
            For i = LBound(words) To UBound(words)
                w = words(i)
                If Len(w) > 0 Then
                    words(i) = UCase$(Left$(w, 1)) & LCase$(Mid$(w, 2))
                End If
            Next i
            ApplyStyleAndAffixes = Join(words, " ")
        Case "sentence"
            If Len(combined) > 0 Then
                ApplyStyleAndAffixes = UCase$(Left$(combined, 1)) & LCase$(Mid$(combined, 2))
            Else
                ApplyStyleAndAffixes = ""
            End If
        Case Else
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
                "Style tidak valid: """ & Style & """."
    End Select
End Function

Private Function InsertThousandsSeparator(ByVal intStr As String, ByVal separator As String) As String
    Dim totalLen As Long
    Dim endPos As Long
    Dim startPos As Long
    Dim chunk As String
    Dim res As String

    totalLen = Len(intStr)
    If Len(separator) = 0 Or totalLen <= 3 Then
        InsertThousandsSeparator = intStr
        Exit Function
    End If

    res = ""
    endPos = totalLen
    Do While endPos > 0
        startPos = endPos - 2
        If startPos < 1 Then startPos = 1
        chunk = Mid$(intStr, startPos, endPos - startPos + 1)
        If Len(res) = 0 Then
            res = chunk
        Else
            res = chunk & separator & res
        End If
        endPos = startPos - 1
    Loop

    InsertThousandsSeparator = res
End Function

Private Function ApplyRupiahAffixes( _
    ByVal numberText As String, _
    ByVal isNegative As Boolean, _
    ByVal Symbol As String, _
    ByVal SpaceAfterSymbol As Boolean, _
    ByVal normNegFormat As String _
) As String
    Dim core As String

    If Len(Symbol) > 0 Then
        If SpaceAfterSymbol Then
            core = Symbol & " " & numberText
        Else
            core = Symbol & numberText
        End If
    Else
        core = numberText
    End If

    If Not isNegative Then
        ApplyRupiahAffixes = core
    ElseIf normNegFormat = "parentheses" Then
        ApplyRupiahAffixes = "(" & core & ")"
    Else
        ApplyRupiahAffixes = "-" & core
    End If
End Function

' ============================================================================
' VALIDATION & STRING UTILITIES
' ============================================================================

Private Function NormalizeRoundingMode(ByVal mode As String) As String
    Dim norm As String
    norm = LCase$(TrimWhitespace(mode))
    If norm <> "half-up" And norm <> "truncate" Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
            "Mode pembulatan tidak valid: """ & mode & """."
    End If
    NormalizeRoundingMode = norm
End Function

Private Sub ParseWholeNumberOption( _
    ByVal optVal As Variant, _
    ByVal minVal As Long, _
    ByVal maxVal As Long, _
    ByVal allowAuto As Boolean, _
    ByVal optionName As String, _
    ByRef OutNumber As Long, _
    ByRef OutIsAuto As Boolean _
)
    Dim vt As Long
    Dim numVal As Double

    OutIsAuto = False
    If IsObject(optVal) Or IsError(optVal) Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
            "Opsi " & optionName & " tidak valid."
    End If

    vt = VarType(optVal)
    If (vt And vbArray) <> 0 Then
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
            "Opsi " & optionName & " tidak valid."
    End If

    If allowAuto And vt = vbString Then
        If LCase$(TrimWhitespace(CStr(optVal))) = "auto" Then
            OutIsAuto = True
            OutNumber = 0
            Exit Sub
        End If
        RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
            "Opsi " & optionName & " harus 'auto' atau integer " & CStr(minVal) & "-" & CStr(maxVal) & "."
    End If

    Select Case vt
        Case vbByte, vbInteger, vbLong, VT_LONGLONG, vbSingle, vbDouble, vbCurrency, vbDecimal
            numVal = CDbl(optVal)
            If numVal <> Int(numVal) Or numVal < minVal Or numVal > maxVal Then
                RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
                    "Opsi " & optionName & " di luar rentang " & CStr(minVal) & "-" & CStr(maxVal) & "."
            End If
            OutNumber = CLng(numVal)
        Case Else
            RaiseTerbilangError TERBILANG_ERR_INVALID_INPUT, "INVALID_INPUT", _
                "Opsi " & optionName & " tidak valid."
    End Select
End Sub

Private Sub AssertMaxDigits(ByVal integerStr As String)
    If Len(integerStr) > MAX_INTEGER_DIGITS Then
        RaiseTerbilangError TERBILANG_ERR_OUT_OF_RANGE, "OUT_OF_RANGE", _
            "Melebihi batas " & CStr(MAX_INTEGER_DIGITS) & " digit."
    End If
End Sub

Private Function IsValidGroupedThousands(ByVal intRaw As String) As Boolean
    Dim parts() As String
    Dim i As Long

    parts = Split(intRaw, ".")
    If UBound(parts) < 1 Then
        IsValidGroupedThousands = False
        Exit Function
    End If

    If Len(parts(0)) < 1 Or Len(parts(0)) > 3 Or Not IsAllAsciiDigits(parts(0)) Then
        IsValidGroupedThousands = False
        Exit Function
    End If

    For i = 1 To UBound(parts)
        If Len(parts(i)) <> 3 Or Not IsAllAsciiDigits(parts(i)) Then
            IsValidGroupedThousands = False
            Exit Function
        End If
    Next i

    IsValidGroupedThousands = True
End Function

Private Function IsAllAsciiDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim code As Long

    If Len(s) = 0 Then
        IsAllAsciiDigits = False
        Exit Function
    End If

    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 48 Or code > 57 Then
            IsAllAsciiDigits = False
            Exit Function
        End If
    Next i

    IsAllAsciiDigits = True
End Function

Private Function IsZeroOnly(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> "0" Then
            IsZeroOnly = False
            Exit Function
        End If
    Next i
    IsZeroOnly = True
End Function

Private Function TrimLeadingZeros(ByVal s As String) As String
    Dim i As Long
    Dim n As Long

    n = Len(s)
    For i = 1 To n
        If Mid$(s, i, 1) <> "0" Then
            TrimLeadingZeros = Mid$(s, i)
            Exit Function
        End If
    Next i
    TrimLeadingZeros = "0"
End Function

Private Function TrimTrailingZeros(ByVal s As String) As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) <> "0" Then
            TrimTrailingZeros = Left$(s, i)
            Exit Function
        End If
    Next i
    TrimTrailingZeros = ""
End Function

Private Function TrimWhitespace(ByVal s As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim n As Long

    n = Len(s)
    If n = 0 Then
        TrimWhitespace = ""
        Exit Function
    End If

    startPos = 1
    Do While startPos <= n
        If Not IsWhitespaceChar(AscW(Mid$(s, startPos, 1))) Then Exit Do
        startPos = startPos + 1
    Loop

    If startPos > n Then
        TrimWhitespace = ""
        Exit Function
    End If

    endPos = n
    Do While endPos >= startPos
        If Not IsWhitespaceChar(AscW(Mid$(s, endPos, 1))) Then Exit Do
        endPos = endPos - 1
    Loop

    TrimWhitespace = Mid$(s, startPos, endPos - startPos + 1)
End Function

Private Function IsWhitespaceChar(ByVal code As Long) As Boolean
    ' Space (32), Tab (9), LF (10), VT (11), FF (12), CR (13), NBSP (160)
    IsWhitespaceChar = (code = 32 Or code = 9 Or code = 10 Or code = 11 Or code = 12 Or code = 13 Or code = 160)
End Function

Private Sub AppendWord(ByRef target As String, ByVal wordStr As String)
    If Len(target) = 0 Then
        target = wordStr
    Else
        target = target & " " & wordStr
    End If
End Sub

Private Function IsWorksheetCall() As Boolean
    On Error Resume Next
    IsWorksheetCall = (TypeName(Application.Caller) = "Range")
    On Error GoTo 0
End Function

Private Function MapErrToExcelError(ByVal errNumber As Long) As Variant
    If errNumber = TERBILANG_ERR_NON_FINITE Or errNumber = TERBILANG_ERR_OUT_OF_RANGE Then
        MapErrToExcelError = CVErr(XL_ERR_NUM)
    Else
        MapErrToExcelError = CVErr(XL_ERR_VALUE)
    End If
End Function

Private Sub RaiseTerbilangError( _
    ByVal errNumber As Long, _
    ByVal errCode As String, _
    ByVal message As String _
)
    Err.Raise errNumber, "IndoTerbilang", "[" & errCode & "] " & message
End Sub
