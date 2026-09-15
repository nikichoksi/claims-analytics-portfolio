Sub BuildLiveDashboard()
    Dim wb As Workbook
    Dim wsData As Worksheet, wsPivot As Worksheet, wsKPI As Worksheet
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim i As Integer
    Dim fieldNames As Variant

    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    Set lo = wsData.ListObjects("ClaimsData")

    ' Re-runnable: remove prior output first (old sheet, and any slicer
    ' caches/slicers left behind by an earlier failed attempt - these cause
    ' "duplicate name" errors on Slicers.Add if not cleared first)
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Pivot_Base").Delete
    Application.DisplayAlerts = True
    Dim scOld As SlicerCache
    For Each scOld In wb.SlicerCaches
        scOld.Delete
    Next scOld
    On Error GoTo 0

    ' ---- Master PivotTable: ClaimStatus x SchoolLevel, Count + Sum ----
    Set wsPivot = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsPivot.Name = "Pivot_Base"

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=lo)
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:="PT_Master")

    With pt
        .PivotFields("ClaimStatus").Orientation = xlRowField
        .PivotFields("SchoolLevel").Orientation = xlColumnField
        With .PivotFields("ClaimId")
            .Orientation = xlDataField
            .Function = xlCount
        End With
        With .PivotFields("IncurredLoss")
            .Orientation = xlDataField
            .Function = xlSum
        End With
    End With

    ' ---- Slicers on Executive_KPIs, connected to this PivotTable ----
    Set wsKPI = wb.Sheets("Executive_KPIs")
    fieldNames = Array("SchoolLevel", "Department", "ClaimStatus")

    ' Anchor to row 40 (well below all existing KPI content) so nothing overlaps
    Dim anchorTop As Double
    anchorTop = wsKPI.Cells(40, 1).Top

    For i = LBound(fieldNames) To UBound(fieldNames)
        Set sc = wb.SlicerCaches.Add2(pt, CStr(fieldNames(i)))
        Set sl = sc.Slicers.Add(wsKPI, , CStr(fieldNames(i)), CStr(fieldNames(i)), _
            anchorTop, wsKPI.Cells(1, 1 + i * 5).Left, 150, 130)
    Next i

    wsKPI.Activate
    MsgBox "Done. Created Pivot_Base with 3 slicers (SchoolLevel, Department, ClaimStatus) on Executive_KPIs."
End Sub
