Attribute VB_Name = "IndoTerbilangExamples"
Option Explicit

' ============================================================================
' Contoh Macro & Playground Excel untuk IndoTerbilang
' Impor file ini bersama src/IndoTerbilang.bas di VBA Editor (Alt + F11).
' ============================================================================

' 1. Contoh cepat di Immediate Window (Ctrl + G di VBA Editor)
Public Sub DemoTerbilangQuickStart()
    Debug.Print "=== indo-terbilang Excel VBA Quick Start ==="
    Debug.Print "Terbilang(1500000): " & Terbilang(1500000)
    Debug.Print "Title + suffix    : " & Terbilang(1500000, "title", "Terbilang:", "rupiah")
    Debug.Print "TerbilangRupiah   : " & TerbilangRupiah(1250.5)
    Debug.Print "Tanpa sen         : " & TerbilangRupiah(1250.6, "sentence", , , , , False)
    Debug.Print "FormatRupiah      : " & FormatRupiah(1500000)
    Debug.Print "Compact           : " & FormatRupiah(1500000, , , , , , , , True)
    Debug.Print "ParseRupiah       : " & ParseRupiah("Rp 1.500.000,50")
    Debug.Print "36 digit (text)   : " & Terbilang("1000000000000000000000000000000000")
End Sub

' 2. Macro interaktif: membuat sheet Demo/Playground Kwitansi & Uji Rumus otomatis
Public Sub BuatSheetDemoIndoTerbilang()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim i As Long
    Dim samples As Variant

    sheetName = "Demo_IndoTerbilang"

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName

    ' Header judul
    ws.Range("A1").Value = "Playground & Demo Rumus indo-terbilang (Excel VBA)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A2").Value = "Catatan: Untuk angka > 15 digit, awali tanda petik satu (') agar tersimpan sebagai teks presisi penuh."
    ws.Range("A2").Font.Italic = True

    ' Header tabel
    ws.Range("A4").Value = "Input Nominal"
    ws.Range("B4").Value = "Terbilang()"
    ws.Range("C4").Value = "TerbilangRupiah() [Title]"
    ws.Range("D4").Value = "FormatRupiah()"
    ws.Range("E4").Value = "FormatRupiah() [Compact]"
    ws.Range("A4:E4").Font.Bold = True
    ws.Range("A4:E4").Interior.Color = RGB(230, 240, 250)

    samples = Array( _
        0, _
        11, _
        1000, _
        1001000, _
        1250.5, _
        1250750.5, _
        -750000, _
        3.14, _
        "12345678901234567890.75" _
    )

    For i = LBound(samples) To UBound(samples)
        Dim r As Long
        r = 5 + i
        If VarType(samples(i)) = vbString Then
            ws.Cells(r, 1).NumberFormat = "@"
        End If
        ws.Cells(r, 1).Value = samples(i)
        ws.Cells(r, 2).Formula = "=Terbilang(A" & r & ")"
        ws.Cells(r, 3).Formula = "=TerbilangRupiah(A" & r & ", ""title"")"
        ws.Cells(r, 4).Formula = "=FormatRupiah(A" & r & ")"
        ws.Cells(r, 5).Formula = "=FormatRupiah(A" & r & ", , , , , , , , TRUE)"
    Next i

    ' Contoh ParseRupiah
    ws.Range("A16").Value = "Uji ParseRupiah"
    ws.Range("A16").Font.Bold = True
    ws.Range("A17").NumberFormat = "@"
    ws.Range("A17").Value = "Rp 1.500.000,50"
    ws.Range("B17").Formula = "=ParseRupiah(A17)"
    ws.Range("C17").Formula = "=TerbilangRupiah(ParseRupiah(A17), ""sentence"")"

    ws.Range("A18").NumberFormat = "@"
    ws.Range("A18").Value = "(Rp 500.000,25)"
    ws.Range("B18").Formula = "=ParseRupiah(A18)"
    ws.Range("C18").Formula = "=TerbilangRupiah(ParseRupiah(A18), ""sentence"")"

    ws.Columns("A:E").AutoFit
    ws.Activate

    MsgBox "Sheet '" & sheetName & "' berhasil dibuat!", vbInformation, "IndoTerbilang"
End Sub

' 3. Macro utilitas: isi kolom di sebelah kanan seleksi dengan TerbilangRupiah (Title Case)
Public Sub IsiKolomTerbilangDariSeleksi()
    Dim cell As Range
    Dim targetCell As Range
    Dim rawVal As Variant
    Dim cleanVal As Variant
    Dim updatedCount As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Pilih terlebih dahulu sel berisi nominal angka.", vbExclamation, "IndoTerbilang"
        Exit Sub
    End If

    updatedCount = 0
    For Each cell In Selection.Cells
        rawVal = cell.Value
        If Not IsEmpty(rawVal) And Not IsError(rawVal) Then
            Set targetCell = cell.Offset(0, 1)
            On Error Resume Next
            If VarType(rawVal) = vbString Then
                If InStr(1, CStr(rawVal), "Rp", vbTextCompare) > 0 _
                    Or InStr(1, CStr(rawVal), "IDR", vbTextCompare) > 0 Then
                    cleanVal = ParseRupiah(rawVal)
                Else
                    cleanVal = rawVal
                End If
            Else
                cleanVal = rawVal
            End If
            targetCell.Value = TerbilangRupiah(cleanVal, "title")
            If Err.Number <> 0 Then
                targetCell.Value = "ERROR: " & Err.Description
                Err.Clear
            Else
                updatedCount = updatedCount + 1
            End If
            On Error GoTo 0
        End If
    Next cell

    MsgBox "Selesai mengisi " & CStr(updatedCount) & " baris terbilang di kolom sebelah kanan.", _
        vbInformation, "IndoTerbilang"
End Sub
