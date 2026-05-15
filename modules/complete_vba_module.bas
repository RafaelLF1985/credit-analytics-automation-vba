Attribute VB_Name = "Módulo1"
Option Explicit

Sub RefreshDashboard()

    Dim wsData As Worksheet
    Dim wsDash As Worksheet
    Dim wsLog As Worksheet

    Set wsData = Sheets("RAW_DATA")
    Set wsDash = Sheets("DASHBOARD")
    Set wsLog = Sheets("AUTOMATION_LOG")

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim totalPortfolio As Double
    Dim approvedVolume As Double
    Dim avgTicket As Double
    Dim approvalRate As Double

    Dim Rejected As Double
    Dim Pending As Double

    Dim spValue As Double
    Dim rjValue As Double
    Dim mgValue As Double
    Dim prValue As Double

    Dim principalRegion As String
    Dim principalRegionValue As Double

    Dim currentVersion As String
    Dim versionNumber As Double

    ' =========================================
    ' KPI CALCULATIONS
    ' =========================================

    totalPortfolio = Application.WorksheetFunction.Sum(wsData.Range("C2:C" & lastRow))

    approvedVolume = Application.WorksheetFunction.SumIf( _
    wsData.Range("D2:D" & lastRow), _
    "Approved", _
    wsData.Range("C2:C" & lastRow))

    Rejected = Application.WorksheetFunction.SumIf( _
    wsData.Range("D2:D" & lastRow), _
    "Rejected", _
    wsData.Range("C2:C" & lastRow))

    Pending = Application.WorksheetFunction.SumIf( _
    wsData.Range("D2:D" & lastRow), _
    "Pending", _
    wsData.Range("C2:C" & lastRow))

    avgTicket = Application.WorksheetFunction.Average(wsData.Range("C2:C" & lastRow))

    approvalRate = Application.WorksheetFunction.CountIf( _
    wsData.Range("D2:D" & lastRow), _
    "Approved") / (lastRow - 1)

    ' =========================================
    ' DASHBOARD KPIs
    ' =========================================

    wsDash.Range("A5").Value = "R$ " & Format(totalPortfolio, "#,##0.00")

    wsDash.Range("C5").Value = "R$ " & Format(approvedVolume, "#,##0.00")

    wsDash.Range("D5").Value = "R$ " & Format(Rejected, "#,##0.00")

    wsDash.Range("E5").Value = "R$ " & Format(Pending, "#,##0.00")

    wsDash.Range("F5").Value = "R$ " & Format(avgTicket, "#,##0.00")

    wsDash.Range("G5").Value = Format(approvalRate, "0.00%")

    ' =========================================
    ' REGION ANALYSIS
    ' =========================================

    spValue = Application.WorksheetFunction.SumIf( _
    wsData.Range("E2:E" & lastRow), _
    "SP", _
    wsData.Range("C2:C" & lastRow))

    rjValue = Application.WorksheetFunction.SumIf( _
    wsData.Range("E2:E" & lastRow), _
    "RJ", _
    wsData.Range("C2:C" & lastRow))

    mgValue = Application.WorksheetFunction.SumIf( _
    wsData.Range("E2:E" & lastRow), _
    "MG", _
    wsData.Range("C2:C" & lastRow))

    prValue = Application.WorksheetFunction.SumIf( _
    wsData.Range("E2:E" & lastRow), _
    "PR", _
    wsData.Range("C2:C" & lastRow))

    wsDash.Range("B10").Value = spValue
    wsDash.Range("B11").Value = rjValue
    wsDash.Range("B12").Value = mgValue
    wsDash.Range("B13").Value = prValue

    ' =========================================
    ' STATUS COUNTERS
    ' =========================================

    wsDash.Range("B16").Value = _
    Application.WorksheetFunction.CountIf( _
    wsData.Range("D2:D" & lastRow), _
    "Approved")

    wsDash.Range("B17").Value = _
    Application.WorksheetFunction.CountIf( _
    wsData.Range("D2:D" & lastRow), _
    "Pending")

    wsDash.Range("B18").Value = _
    Application.WorksheetFunction.CountIf( _
    wsData.Range("D2:D" & lastRow), _
    "Rejected")

    ' =========================================
    ' PRINCIPAL REGION
    ' =========================================

    principalRegion = "SP"
    principalRegionValue = spValue

    If rjValue > principalRegionValue Then
        principalRegion = "RJ"
        principalRegionValue = rjValue
    End If

    If mgValue > principalRegionValue Then
        principalRegion = "MG"
        principalRegionValue = mgValue
    End If

    If prValue > principalRegionValue Then
        principalRegion = "PR"
        principalRegionValue = prValue
    End If

    ' =========================================
    ' AUTOMATION LOG
    ' =========================================

    wsLog.Range("B2").Value = "NOT"

    wsLog.Range("B3").Value = "NOT"

    wsLog.Range("B5").Value = principalRegion

    wsLog.Range("B7").Value = Environ("Username")

    wsLog.Range("B8").Value = Application.OperatingSystem

    wsLog.Range("B9").Value = Now
    wsLog.Range("B9").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    currentVersion = wsLog.Range("B12").Value

    If currentVersion = "" Then

        versionNumber = 1#

    Else

        currentVersion = Replace(currentVersion, "v", "")

        versionNumber = Val(currentVersion)

        versionNumber = versionNumber + 0.1

    End If

    wsLog.Range("B12").Value = "v" & Format(versionNumber, "0.0")

    wsLog.Range("B13").Value = Date
    wsLog.Range("B13").NumberFormat = "dd/mm/yyyy"

    ' =========================================
    ' VISUAL FORMAT
    ' =========================================

    With wsDash.Range("A5:G5")

        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(15, 15, 15)
        .Font.Color = RGB(255, 255, 255)

    End With

    Call ExportPDF

End Sub

Sub ExportPDF()

    Dim wsLog As Worksheet
    Dim wsDash As Worksheet

    Set wsLog = Sheets("AUTOMATION_LOG")
    Set wsDash = Sheets("DASHBOARD")

    ' =========================================
    ' PDF SETTINGS
    ' =========================================

    With wsDash.PageSetup

        .Orientation = xlLandscape

        .Zoom = False

        .FitToPagesWide = 1
        .FitToPagesTall = 1

        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)

        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)

        .CenterHorizontally = True
        .CenterVertically = True

    End With

    ' =========================================
    ' EXPORT PDF
    ' =========================================

    wsDash.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & "\Executive_Dashboard.pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    wsLog.Range("B2").Value = "YES"

    wsLog.Range("B9").Value = Now
    wsLog.Range("B9").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    Call ExportCSV

End Sub

Sub ExportCSV()

    Dim wsLog As Worksheet

    Set wsLog = Sheets("AUTOMATION_LOG")

    Sheets("RAW_DATA").Copy

    ActiveWorkbook.SaveAs _
        Filename:=ThisWorkbook.Path & "\RAW_DATA_EXPORT.csv", _
        FileFormat:=xlCSV

    ActiveWorkbook.Close False

    wsLog.Range("B3").Value = "YES"

    wsLog.Range("B9").Value = Now
    wsLog.Range("B9").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    Call ExportLOG

End Sub

Sub ExportLOG()

    Dim wsLog As Worksheet

    Set wsLog = Sheets("AUTOMATION_LOG")

    wsLog.Copy

    ActiveWorkbook.SaveAs _
        Filename:=ThisWorkbook.Path & "\AUTOMATION_LOG_EXPORT.csv", _
        FileFormat:=xlCSV

    ActiveWorkbook.Close False

    MsgBox "Dashboard updated successfully! LOG|PDF|CSV baixados!", vbInformation

End Sub

