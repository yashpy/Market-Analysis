' ============================================================
' Temple Allen Industries – Business Development Report Macro
' Author: Yadnesh Deshpande
' Description: Automates weekly BD status report from raw data
' ============================================================

Option Explicit

' ── Main Entry Point ─────────────────────────────────────────
Sub GenerateBDReport()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsData   As Worksheet
    Dim wsReport As Worksheet

    ' Get or create sheets
    Set wsData = GetOrCreateSheet("Raw_Data")
    Set wsReport = GetOrCreateSheet("BD_Report")

    ' Seed sample data if empty
    If wsData.Cells(2, 1).Value = "" Then Call SeedSampleData(wsData)

    ' Build report
    Call BuildReportHeader(wsReport)
    Call SummarizeByStatus(wsData, wsReport)
    Call SummarizeByIndustry(wsData, wsReport)
    Call FormatReport(wsReport)

    wsReport.Activate
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "BD Report generated successfully!", vbInformation, "TAI Report"
End Sub

' ── Seed Sample Proposals Data ────────────────────────────────
Sub SeedSampleData(ws As Worksheet)
    Dim headers() As String
    headers = Split("Customer,Industry,Status,Units_Quoted,Quote_Value,Rep,Date", ",")

    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Cells(1, i + 1).Interior.Color = RGB(31, 56, 100)
        ws.Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
    Next i

    Dim data As Variant
    data = Array( _
        Array("Boeing MRO",         "Aerospace", "Won",         2, 700000, "Y. Deshpande", "2026-01-15"), _
        Array("Lockheed Martin",    "Defense",   "Negotiating", 1, 350000, "Y. Deshpande", "2026-02-01"), _
        Array("Vestas Wind",        "Windpower", "Sent",        1, 350000, "Y. Deshpande", "2026-02-20"), _
        Array("Huntington Ingalls", "Marine",    "Won",         1, 350000, "Y. Deshpande", "2026-03-05"), _
        Array("Delta TechOps",      "Aerospace", "Draft",       3, 1050000,"Y. Deshpande", "2026-03-18"), _
        Array("Airbus MRO",         "Aerospace", "Lost",        2, 700000, "Y. Deshpande", "2026-01-30"), _
        Array("US Navy",            "Defense",   "Negotiating", 4, 1400000,"Y. Deshpande", "2026-04-01") _
    )

    Dim r As Integer
    For r = 0 To UBound(data)
        Dim j As Integer
        For j = 0 To 6
            ws.Cells(r + 2, j + 1).Value = data(r)(j)
        Next j
        ' Format currency
        ws.Cells(r + 2, 5).NumberFormat = "$#,##0"
    Next r

    ws.Columns("A:G").AutoFit
End Sub

' ── Report Header ─────────────────────────────────────────────
Sub BuildReportHeader(ws As Worksheet)
    ws.Cells.Clear

    With ws.Range("A1:F1")
        .Merge
        .Value = "Temple Allen Industries – BD Pipeline Report"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    With ws.Range("A2:F2")
        .Merge
        .Value = "Generated: " & Format(Now(), "MMMM DD, YYYY  HH:MM") & "   |   Author: Yadnesh Deshpande"
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
    End With
End Sub

' ── Summarize by Status ───────────────────────────────────────
Sub SummarizeByStatus(wsData As Worksheet, wsReport As Worksheet)
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' Section header
    Dim startRow As Long: startRow = 4
    With wsReport.Range("A" & startRow & ":F" & startRow)
        .Merge
        .Value = "PIPELINE SUMMARY BY STATUS"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 117, 182)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With

    ' Column headers
    Dim cols() As String
    cols = Split("Status,# Deals,Units Quoted,Total Value,Avg Deal Size,Win Rate %", ",")
    Dim c As Integer
    For c = 0 To 5
        wsReport.Cells(startRow + 1, c + 1).Value = cols(c)
        wsReport.Cells(startRow + 1, c + 1).Font.Bold = True
        wsReport.Cells(startRow + 1, c + 1).Interior.Color = RGB(217, 225, 242)
    Next c

    Dim statuses() As String
    statuses = Split("Won,Negotiating,Sent,Draft,Lost", ",")

    Dim totalWon As Double: totalWon = 0
    Dim totalDeals As Long: totalDeals = 0
    Dim wonDeals As Long: wonDeals = 0

    Dim s As Integer
    For s = 0 To UBound(statuses)
        Dim r As Long
        Dim dealCount As Long:   dealCount = 0
        Dim unitCount As Long:   unitCount = 0
        Dim valueSum As Double:  valueSum = 0

        For r = 2 To lastRow
            If wsData.Cells(r, 3).Value = statuses(s) Then
                dealCount = dealCount + 1
                unitCount = unitCount + wsData.Cells(r, 4).Value
                valueSum  = valueSum  + wsData.Cells(r, 5).Value
            End If
        Next r

        totalDeals = totalDeals + dealCount
        If statuses(s) = "Won" Then
            wonDeals = dealCount
            totalWon = valueSum
        End If

        Dim dataRow As Long: dataRow = startRow + 2 + s
        Dim avgDeal As Double
        avgDeal = IIf(dealCount > 0, valueSum / dealCount, 0)

        wsReport.Cells(dataRow, 1).Value = statuses(s)
        wsReport.Cells(dataRow, 2).Value = dealCount
        wsReport.Cells(dataRow, 3).Value = unitCount
        wsReport.Cells(dataRow, 4).Value = valueSum
        wsReport.Cells(dataRow, 4).NumberFormat = "$#,##0"
        wsReport.Cells(dataRow, 5).Value = avgDeal
        wsReport.Cells(dataRow, 5).NumberFormat = "$#,##0"
        wsReport.Cells(dataRow, 6).Value = IIf(totalDeals > 0, dealCount / totalDeals, 0)
        wsReport.Cells(dataRow, 6).NumberFormat = "0.0%"

        If statuses(s) = "Won" Then
            wsReport.Range(wsReport.Cells(dataRow, 1), wsReport.Cells(dataRow, 6)).Interior.Color = RGB(226, 239, 218)
        ElseIf statuses(s) = "Lost" Then
            wsReport.Range(wsReport.Cells(dataRow, 1), wsReport.Cells(dataRow, 6)).Interior.Color = RGB(255, 235, 235)
        End If
    Next s
End Sub

' ── Summarize by Industry ─────────────────────────────────────
Sub SummarizeByIndustry(wsData As Worksheet, wsReport As Worksheet)
    Dim startRow As Long: startRow = 13

    With wsReport.Range("A" & startRow & ":F" & startRow)
        .Merge
        .Value = "PIPELINE SUMMARY BY INDUSTRY"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 117, 182)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With

    Dim cols() As String
    cols = Split("Industry,# Deals,Units Quoted,Total Value,Won Value,Win %", ",")
    Dim c As Integer
    For c = 0 To 5
        wsReport.Cells(startRow + 1, c + 1).Value = cols(c)
        wsReport.Cells(startRow + 1, c + 1).Font.Bold = True
        wsReport.Cells(startRow + 1, c + 1).Interior.Color = RGB(217, 225, 242)
    Next c

    Dim industries() As String
    industries = Split("Aerospace,Defense,Marine,Windpower", ",")

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim ind As Integer
    For ind = 0 To UBound(industries)
        Dim r As Long
        Dim deals As Long: deals = 0
        Dim units As Long: units = 0
        Dim total As Double: total = 0
        Dim won As Double: won = 0

        For r = 2 To lastRow
            If wsData.Cells(r, 2).Value = industries(ind) Then
                deals = deals + 1
                units = units + wsData.Cells(r, 4).Value
                total = total + wsData.Cells(r, 5).Value
                If wsData.Cells(r, 3).Value = "Won" Then won = won + wsData.Cells(r, 5).Value
            End If
        Next r

        Dim dataRow As Long: dataRow = startRow + 2 + ind
        wsReport.Cells(dataRow, 1).Value = industries(ind)
        wsReport.Cells(dataRow, 2).Value = deals
        wsReport.Cells(dataRow, 3).Value = units
        wsReport.Cells(dataRow, 4).Value = total
        wsReport.Cells(dataRow, 4).NumberFormat = "$#,##0"
        wsReport.Cells(dataRow, 5).Value = won
        wsReport.Cells(dataRow, 5).NumberFormat = "$#,##0"
        wsReport.Cells(dataRow, 6).Value = IIf(total > 0, won / total, 0)
        wsReport.Cells(dataRow, 6).NumberFormat = "0.0%"

        If ind Mod 2 = 0 Then
            wsReport.Range(wsReport.Cells(dataRow, 1), wsReport.Cells(dataRow, 6)).Interior.Color = RGB(242, 242, 242)
        End If
    Next ind
End Sub

' ── Format Report ─────────────────────────────────────────────
Sub FormatReport(ws As Worksheet)
    Dim c As Integer
    Dim widths() As Integer
    widths = Array(20, 10, 14, 14, 14, 12)
    For c = 1 To 6
        ws.Columns(c).ColumnWidth = widths(c - 1)
    Next c

    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Columns("A:F").HorizontalAlignment = xlCenter
    ws.Columns("A").HorizontalAlignment = xlLeft
End Sub

' ── Helper: Get or Create Sheet ───────────────────────────────
Function GetOrCreateSheet(name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = name
    End If
    Set GetOrCreateSheet = ws
End Function
