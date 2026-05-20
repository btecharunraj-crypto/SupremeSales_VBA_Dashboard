Attribute VB_Name = "Module1"
Sub muralibp()

'Change your Pending file name
    Dim SourceBook As Workbook
    Dim SourceSheet As Worksheet
    Dim PenSheet As Worksheet
    Dim FolderPath As String
    Dim TodaysDate As String
    Dim FileName1 As String
    Dim FileName2 As String
    Dim FullFilePath1 As String
    Dim FullFilePath2 As String
    Dim FileSystem1 As Object
    Dim FileSystem2 As Object
    
    FolderPath = "D:\PRIYANKA 1 PVC FILES\Macro_Projects\Project1\"
    TodaysDate = Format(Date, "DDMMYY")
    FileName1 = "pending" & TodaysDate & ".XLSX"
    FullFilePath1 = FolderPath & FileName1
    Set FileSystem1 = CreateObject("Scripting.FileSystemObject")
    
    If FileSystem1.FileExists(FullFilePath1) Then
        Set SourceBook = Workbooks.Open(FullFilePath1)
        MsgBox "File opened successfully: " & FullFilePath1
    Else
        MsgBox "File not found: " & FullFilePath1
    End If
    
    Set SourceSheet = SourceBook.Worksheets("Sheet1")
    SourceSheet.Activate
    Dim LrSou As Long
    Dim RngMgrSource As Range
  
    LrSou = SourceSheet.Cells(SourceSheet.Rows.Count, "M").End(xlUp).Row
    Set RngMgrSource = SourceSheet.Range("M1:M" & LrSou)
    RngMgrSource.AutoFilter Field:=1, Criteria1:="SD0045", Operator:=xlFilterValues
    SourceSheet.Range("A1").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Application.DisplayAlerts = True
    SourceSheet.AutoFilterMode = False

'Pasted to Pendingsheet
    Set PenSheet = ThisWorkbook.Worksheets("Pending")
    PenSheet.Activate
    PenSheet.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

'Pendingsheet Filtering and removing
    Dim LrPen As Long
    Dim RngPdescPen As Range
    Dim RngMatPen As Range
    Dim LcHeadPen As Long
    LrPen = PenSheet.Cells(PenSheet.Rows.Count, "N").End(xlUp).Row
    Set RngPdescPen = PenSheet.Range("N1:N" & LrPen)
    Set RngMatPen = PenSheet.Range("G1:G" & LrPen)
    RngPdescPen.AutoFilter Field:=1, Criteria1:=Array("P-D", "P-F", "F-F", "01", "S11", "E-H", "C-P", "M-F", "CS-P", "P-T"), Operator:=xlFilterValues
    PenSheet.Range("A2").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    PenSheet.AutoFilterMode = False
    RngMatPen.AutoFilter Field:=1, Criteria1:="*Sample*", Operator:=xlFilterValues
    PenSheet.Range("A2").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    PenSheet.AutoFilterMode = False
    LcHeadPen = PenSheet.Cells(1, PenSheet.Columns.Count).End(xlToLeft).Column
    PenSheet.Cells(1, LcHeadPen + 1) = "Pending_Liters"
    PenSheet.Cells(1, LcHeadPen + 2) = "Pending_Actual_Liters"

'Pendingsheet New columns working(Vlookup with Sumifsheet)
    Dim VBook As Workbook
    Dim VSheet As Worksheet
    Dim LrVs As Long
    Dim RngVs As Range
    Dim iPen As Long
    Dim LvalPen As String
    Dim LitersPen As Variant
    Dim LitersAmtPen As Double
    Set VBook = Workbooks.Open("D:\PRIYANKA 1 PVC FILES\Macro_Projects\Project1\sumif.xlsx")
    Set VSheet = VBook.Worksheets("Sheet1")
    LrVs = VSheet.Cells(VSheet.Rows.Count, "A").End(xlUp).Row
    Set RngVs = VSheet.Range("B1:D" & LrVs)
    LrPen = PenSheet.Cells(PenSheet.Rows.Count, "F").End(xlUp).Row
    For iPen = 2 To LrPen
    LvalPen = PenSheet.Cells(iPen, "F").Value
    LitersPen = Application.VLookup(LvalPen, RngVs, 3, False)
    If Not IsError(LitersPen) Then
    PenSheet.Cells(iPen, "Q").Value = LitersPen
    Else
    PenSheet.Cells(iPen, "Q").Value = 0
    End If

'Pending file Setting formulas for Liters and Literssales
    LitersAmtPen = PenSheet.Cells(iPen, "K").Value * PenSheet.Cells(iPen, "Q").Value
    PenSheet.Cells(iPen, "R") = LitersAmtPen
    Next iPen
    PenSheet.Columns.AutoFit
    PenSheet.Activate
    PenSheet.Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Selection.Interior.Color = vbYellow

'Change your Sales file name
    Dim SaleBook As Workbook
    Dim SaleSheet As Worksheet
    Dim SSheet As Worksheet
    
    FileName2 = "Salesdetail" & TodaysDate & ".XLSX"
    FullFilePath2 = FolderPath & FileName2
    Set FileSystem2 = CreateObject("Scripting.FileSystemObject")
    
    If FileSystem2.FileExists(FullFilePath2) Then
        Set SaleBook = Workbooks.Open(FullFilePath2)
        MsgBox "File opened successfully: " & FullFilePath2
    Else
        MsgBox "File not found: " & FullFilePath2
    End If
    
    Set SaleSheet = SaleBook.Worksheets("Sheet1")
    SaleSheet.Activate
    Dim LrSou1 As Long
    Dim RngMgrSource1 As Range
    LrSou1 = SaleSheet.Cells(SaleSheet.Rows.Count, "G").End(xlUp).Row
    Set RngMgrSource1 = SaleSheet.Range("G1:G" & LrSou1)
    RngMgrSource1.AutoFilter Field:=1, Criteria1:="TAMILNADU (MURALI)", Operator:=xlFilterValues
    SaleSheet.Range("A1").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Application.DisplayAlerts = True
    SourceSheet.AutoFilterMode = False
  
'Pasted to Salessheet
    Set SSheet = ThisWorkbook.Worksheets("SalesData")
    SSheet.Activate
    SSheet.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
'Salessheet Filtering and removing
    Dim LrSal As Long
    Dim RngPdescSal As Range
    Dim RngMatSal As Range
    Dim LcSalHead As Long
    LrSal = SSheet.Cells(SSheet.Rows.Count, "H").End(xlUp).Row
    Set RngPdescSal = SSheet.Range("H1:H" & LrSal)
    Set RngMatSal = SSheet.Range("J1:J" & LrSal)
    RngPdescSal.AutoFilter Field:=1, Criteria1:=Array("P-D", "P-F", "F-F", "01", "S11", "E-H", "C-P", "M-F", "CS-P", "P-T"), Operator:=xlFilterValues
    SSheet.Range("A2").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    SSheet.AutoFilterMode = False
    RngMatSal.AutoFilter Field:=1, Criteria1:="*Sample*", Operator:=xlFilterValues
    SSheet.Range("A2").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    SSheet.AutoFilterMode = False
    LcSalHead = SSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    SSheet.Cells(1, LcSalHead + 1).Value = "Sales_Liters"
    SSheet.Cells(1, LcSalHead + 2).Value = "Sales_Actual_Liters"
    SSheet.Cells(1, LcSalHead + 3).Value = "SalesAmount"
  
'Salessheet New columns working(Vlookup with Sumifsheet)
    Dim iSal As Long
    Dim LvalSal As String
    Dim LitersSal As Variant
    Dim LitersAmtSal As Double
    Dim BillAmtSal As Double
    LrSal = SSheet.Cells(SSheet.Rows.Count, "I").End(xlUp).Row
    For iSal = 2 To LrSal
    LvalSal = SSheet.Cells(iSal, "I").Value
    LitersSal = Application.VLookup(LvalSal, RngVs, 3, False)
    If Not IsError(LitersSal) Then
    SSheet.Cells(iSal, "AC").Value = LitersSal
    Else
    SSheet.Cells(iSal, "AC").Value = 0
    End If
    
'Sales file Setting formulas for Liters and Literssales
    LitersAmtSal = SSheet.Cells(iSal, "AC").Value * SSheet.Cells(iSal, "N").Value
    SSheet.Cells(iSal, "AD") = LitersAmtSal
    BillAmtSal = SSheet.Cells(iSal, "AA").Value - SSheet.Cells(iSal, "X").Value - SSheet.Cells(iSal, "Y").Value _
    - SSheet.Cells(iSal, "Z").Value
    SSheet.Cells(iSal, "AE") = BillAmtSal
    Next iSal
    SSheet.Columns.AutoFit
    SSheet.Activate
    SSheet.Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Selection.Interior.Color = vbRed
    Selection.Font.Color = vbWhite
    
'All table Duplicates working
    Dim PiSheet As Worksheet
    Dim LrPi1F As Long
    Dim LrPi1S As Long
    Dim LrPi1S3 As Long
    Dim RngPi1S1 As Range
    Dim RngPi1FS As Range
    Dim RngPi1Full As Range
    Set PiSheet = ThisWorkbook.Worksheets("Pivot")
    PenSheet.Activate
    PenSheet.Range("A1:N" & LrPen).Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).Copy
    PiSheet.Activate
    PiSheet.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    SSheet.Activate
    SSheet.Range("D1:H" & LrSal).Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).Copy
    PiSheet.Activate
    LrPi1F = PiSheet.Cells(PiSheet.Rows.Count, "N").End(xlUp).Row
    PiSheet.Range("A" & LrPi1F + 2).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    LrPi1S = PiSheet.Cells(PiSheet.Rows.Count, "E").End(xlUp).Row
    Set RngPi1S1 = PiSheet.Range("E" & LrPi1F + 2 & ":E" & LrPi1S)
    RngPi1S1.Copy
    PiSheet.Range("N" & LrPi1F + 2).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    PiSheet.Range("F" & LrPi1F + 2 & ":M" & LrPi1F + 2).Value = "NA"
    PiSheet.Range("A" & LrPi1F + 1 & ":A" & LrPi1F + 2).EntireRow.Delete
    LrPi1S = PiSheet.Cells(PiSheet.Rows.Count, "N").End(xlUp).Row
    Set RngPi1FS = PiSheet.Range("N1:N" & LrPi1S)
    RngPi1FS.AutoFilter Field:=1, Criteria1:=Array("BMT", "R-T", "R-S"), Operator:=xlFilterValues
    PiSheet.Range("A2").Select
    Application.DisplayAlerts = False
    Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Delete
    Application.DisplayAlerts = True
    PiSheet.AutoFilterMode = False
    PiSheet.Columns("C:N").Delete
    LrPi1S = PiSheet.Cells(PiSheet.Rows.Count, "A").End(xlUp).Row
    Set RngPi1Full = PiSheet.Range("A1:B" & LrPi1S)
    RngPi1Full.RemoveDuplicates Columns:=1, Header:=xlNo
    
'All table sumifs working
       Dim TotalRevB As Double
       Dim Customer1 As String
       Dim i1 As Long
       Dim LrPi1 As Long
       Dim TotalB11 As Double
       Dim TotalB12 As Double
       Dim TotalB13 As Double
       Dim TotalB14 As Double
       Dim TotalB15 As Double
       Dim GTotal1 As Double
       Dim TS1 As Double
       LrPi1 = PiSheet.Cells(PiSheet.Rows.Count, "A").End(xlUp).Row
       For i1 = 2 To LrPi1
    Customer1 = PiSheet.Cells(i1, "A").Value
    TotalRevB = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer1, _
    SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB11 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer1, _
    SSheet.Range("H1:H" & LrSal), "BMT", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB12 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer1, _
    SSheet.Range("H1:H" & LrSal), "R-T", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB13 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer1, _
    SSheet.Range("H1:H" & LrSal), "R-S", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB14 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer1, _
    SSheet.Range("H1:H" & LrSal), "01", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB15 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer1, _
    SSheet.Range("H1:H" & LrSal), "S11", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
     PiSheet.Cells(i1, "C").Value = (TotalRevB - TotalB11 - TotalB12 - TotalB13 - TotalB14 - TotalB15) / (10 ^ 5)
       Dim TotalRevP As Double
       Dim TotalP11 As Double
       Dim TotalP12 As Double
       Dim TotalP13 As Double
       Dim TotalP14 As Double
       Dim TotalP15 As Double
    TotalRevP = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer1, _
    PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP11 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer1, _
    PenSheet.Range("N1:N" & LrPen), "BMT", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP12 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer1, _
    PenSheet.Range("N1:N" & LrPen), "R-T", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP13 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer1, _
    PenSheet.Range("N1:N" & LrPen), "R-S", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP14 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer1, _
    PenSheet.Range("N1:N" & LrPen), "01", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP15 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer1, _
    PenSheet.Range("N1:N" & LrPen), "S11", PenSheet.Range("M1:M" & LrPen), "SD0045")
    PiSheet.Cells(i1, "D").Value = (TotalRevP - TotalP11 - TotalP12 - TotalP13 - TotalP14 - TotalP15) / (10 ^ 5)
    GTotal1 = PiSheet.Cells(i1, "C").Value + PiSheet.Cells(i1, "D")
    PiSheet.Cells(i1, "E") = GTotal1
    Next i1
    
'MF table Duplicates working
    PenSheet.Activate
    RngPdescPen.AutoFilter Field:=1, Criteria1:=Array("M-G", "M-R", "M-S"), Operator:=xlFilterValues
    PenSheet.Range("A1:B" & LrPen).SpecialCells(xlCellTypeVisible).Copy
    Application.DisplayAlerts = False
    PiSheet.Range("G1").PasteSpecial xlPasteValues
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    PenSheet.AutoFilterMode = False
    Dim LrPi2f As Long
    Dim LrPi2S As Long
    LrPi2f = PiSheet.Cells(PiSheet.Rows.Count, "G").End(xlUp).Row
    SSheet.Activate
    RngPdescSal.AutoFilter Field:=1, Criteria1:=Array("M-G", "M-R", "M-S"), Operator:=xlFilterValues
    SSheet.Range("D1:E" & LrSal).SpecialCells(xlCellTypeVisible).Copy
    Application.DisplayAlerts = False
    PiSheet.Range("G" & LrPi2f + 1).PasteSpecial xlPasteValues
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    SSheet.AutoFilterMode = False
    LrPi2S = PiSheet.Cells(PiSheet.Rows.Count, "G").End(xlUp).Row
    PiSheet.Range("G" & LrPi2f + 2 & ":H" & LrPi2S).Cut Destination:=PiSheet.Range("G" & LrPi2f + 1)
    Application.CutCopyMode = False
    Dim RngPi2Full As Range
    Set RngPi2Full = PiSheet.Range("G1:H" & LrPi2S - 1)
    RngPi2Full.RemoveDuplicates Columns:=1, Header:=xlNo

     
'MF table sumifs working
       Dim Customer2 As String
       Dim i2 As Long
       Dim LrPi2 As Long
       Dim TotalB21 As Double
       Dim TotalB22 As Double
       Dim TotalB23 As Double
       Dim GTotal2 As Double
       LrPi2 = PiSheet.Cells(PiSheet.Rows.Count, "G").End(xlUp).Row
         For i2 = 2 To LrPi2
    Customer2 = PiSheet.Cells(i2, "G").Value
    TotalB21 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer2, _
    SSheet.Range("H1:H" & LrSal), "M-G", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB22 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer2, _
    SSheet.Range("H1:H" & LrSal), "M-R", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB23 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer2, _
    SSheet.Range("H1:H" & LrSal), "M-S", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    PiSheet.Cells(i2, "I").Value = (TotalB21 + TotalB22 + TotalB23) / (10 ^ 5)
       Dim TotalP21 As Double
       Dim TotalP22 As Double
       Dim TotalP23 As Double
    TotalP21 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer2, _
    PenSheet.Range("N1:N" & LrPen), "M-G", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP22 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer2, _
    PenSheet.Range("N1:N" & LrPen), "M-R", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP23 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer2, _
    PenSheet.Range("N1:N" & LrPen), "M-S", PenSheet.Range("M1:M" & LrPen), "SD0045")
    PiSheet.Cells(i2, "J").Value = (TotalP21 + TotalP22 + TotalP23) / (10 ^ 5)
    GTotal2 = PiSheet.Cells(i2, "I").Value + PiSheet.Cells(i2, "J")
    PiSheet.Cells(i2, "K") = GTotal2
    Next i2
 
'CPVC table Duplicates working
    PenSheet.Activate
    RngPdescPen.AutoFilter Field:=1, Criteria1:=Array("P-C", "M-C", "F-C"), Operator:=xlFilterValues
    PenSheet.Range("A1:B" & LrPen).SpecialCells(xlCellTypeVisible).Copy
    Application.DisplayAlerts = False
    PiSheet.Range("M1").PasteSpecial xlPasteValues
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    PenSheet.AutoFilterMode = False
    Dim LrPi3F As Long
    Dim LrPi3S As Long
    LrPi3F = PiSheet.Cells(PiSheet.Rows.Count, "M").End(xlUp).Row
    SSheet.Activate
    RngPdescSal.AutoFilter Field:=1, Criteria1:=Array("P-C", "M-C", "F-C"), Operator:=xlFilterValues
    SSheet.Range("D1:E" & LrSal).SpecialCells(xlCellTypeVisible).Copy
    Application.DisplayAlerts = False
    PiSheet.Range("M" & LrPi3F + 1).PasteSpecial xlPasteValues
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    SSheet.AutoFilterMode = False
    LrPi3S = PiSheet.Cells(PiSheet.Rows.Count, "M").End(xlUp).Row
    PiSheet.Range("M" & LrPi3F + 2 & ":N" & LrPi3S).Cut Destination:=PiSheet.Range("M" & LrPi3F + 1)
    Application.CutCopyMode = False
    Dim RngPi3Full As Range
    Set RngPi3Full = PiSheet.Range("M1:N" & LrPi3S - 1)
    RngPi3Full.RemoveDuplicates Columns:=1, Header:=xlNo
   
     
'CPVC table sumifs working
       Dim Customer3 As String
       Dim i3 As Long
       Dim LrPi3 As Long
       Dim TotalB31 As Double
       Dim TotalB32 As Double
       Dim TotalB33 As Double
       Dim GTotal3 As Double
       LrPi3 = PiSheet.Cells(PiSheet.Rows.Count, "M").End(xlUp).Row
       For i3 = 2 To LrPi3
    Customer3 = PiSheet.Cells(i3, "M").Value
    TotalB31 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer3, _
    SSheet.Range("H1:H" & LrSal), "P-C", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB32 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer3, _
    SSheet.Range("H1:H" & LrSal), "M-C", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB33 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer3, _
    SSheet.Range("H1:H" & LrSal), "F-C", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    PiSheet.Cells(i3, "O").Value = (TotalB31 + TotalB32 + TotalB33) / (10 ^ 5)
       Dim TotalP31 As Double
       Dim TotalP32 As Double
       Dim TotalP33 As Double
    TotalP31 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer3, _
    PenSheet.Range("N1:N" & LrPen), "P-C", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP32 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer3, _
    PenSheet.Range("N1:N" & LrPen), "M-C", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP33 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer3, _
    PenSheet.Range("N1:N" & LrPen), "F-C", PenSheet.Range("M1:M" & LrPen), "SD0045")
    PiSheet.Cells(i3, "P").Value = (TotalP31 + TotalP32 + TotalP33) / (10 ^ 5)
    GTotal3 = PiSheet.Cells(i3, "O").Value + PiSheet.Cells(i3, "P")
    PiSheet.Cells(i3, "Q") = GTotal3
    Next i3
 
'Tanks table Duplicates working
    PenSheet.Activate
    RngPdescPen.AutoFilter Field:=1, Criteria1:=Array("BMT", "R-T", "R-S"), Operator:=xlFilterValues
    PenSheet.Range("A1:B" & LrPen).SpecialCells(xlCellTypeVisible).Copy
    Application.DisplayAlerts = False
    PiSheet.Range("S1").PasteSpecial xlPasteValues
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    PenSheet.AutoFilterMode = False
    Dim LrPi4F As Long
    Dim LrPi4S As Long
    LrPi4F = PiSheet.Cells(PiSheet.Rows.Count, "S").End(xlUp).Row
    SSheet.Activate
    RngPdescSal.AutoFilter Field:=1, Criteria1:=Array("BMT", "R-T", "R-S"), Operator:=xlFilterValues
    SSheet.Range("D1:E" & LrSal).SpecialCells(xlCellTypeVisible).Copy
    Application.DisplayAlerts = False
    PiSheet.Range("S" & LrPi4F + 1).PasteSpecial xlPasteValues
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    SSheet.AutoFilterMode = False
    LrPi4S = PiSheet.Cells(PiSheet.Rows.Count, "S").End(xlUp).Row
    PiSheet.Range("S" & LrPi4F + 2 & ":T" & LrPi4S).Cut Destination:=PiSheet.Range("S" & LrPi4F + 1)
    Application.CutCopyMode = False
    Dim RngPi4Full As Range
    Set RngPi4Full = PiSheet.Range("S1:T" & LrPi4S - 1)
    RngPi4Full.RemoveDuplicates Columns:=1, Header:=xlNo
    
'Tanks table sumifs working
       Dim Customer4 As String
       Dim i4 As Long
       Dim LrPi4 As Long
       Dim TotalB41 As Double
       Dim TotalB42 As Double
       Dim TotalB43 As Double
       Dim TotalB44 As Double
       Dim TotalB45 As Double
       Dim TotalB46 As Double
       LrPi4 = PiSheet.Cells(PiSheet.Rows.Count, "S").End(xlUp).Row
       For i4 = 2 To LrPi4
    Customer4 = PiSheet.Cells(i4, "S").Value
    TotalB41 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer4, _
    SSheet.Range("H1:H" & LrSal), "BMT", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB42 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer4, _
    SSheet.Range("H1:H" & LrSal), "R-T", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB43 = Application.SumIfs(SSheet.Range("AE1:AE" & LrSal), SSheet.Range("D1:D" & LrSal), Customer4, _
    SSheet.Range("H1:H" & LrSal), "R-S", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    PiSheet.Cells(i4, "U").Value = (TotalB41 + TotalB42 + TotalB43) / (10 ^ 5)
    TotalB44 = Application.SumIfs(SSheet.Range("AD1:AD" & LrSal), SSheet.Range("D1:D" & LrSal), Customer4, _
    SSheet.Range("H1:H" & LrSal), "BMT", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB45 = Application.SumIfs(SSheet.Range("AD1:AD" & LrSal), SSheet.Range("D1:D" & LrSal), Customer4, _
    SSheet.Range("H1:H" & LrSal), "R-T", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    TotalB46 = Application.SumIfs(SSheet.Range("AD1:AD" & LrSal), SSheet.Range("D1:D" & LrSal), Customer4, _
    SSheet.Range("H1:H" & LrSal), "R-S", SSheet.Range("G1:G" & LrSal), "TAMILNADU (MURALI)")
    PiSheet.Cells(i4, "V").Value = (TotalB44 + TotalB45 + TotalB46)
       Dim TotalP41 As Double
       Dim TotalP42 As Double
       Dim TotalP43 As Double
       Dim TotalP44 As Double
       Dim TotalP45 As Double
       Dim TotalP46 As Double
    TotalP41 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer4, _
    PenSheet.Range("N1:N" & LrPen), "BMT", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP42 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer4, _
    PenSheet.Range("N1:N" & LrPen), "R-T", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP43 = Application.SumIfs(PenSheet.Range("L1:L" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer4, _
    PenSheet.Range("N1:N" & LrPen), "R-S", PenSheet.Range("M1:M" & LrPen), "SD0045")
    PiSheet.Cells(i4, "W").Value = (TotalP41 + TotalP42 + TotalP43) / (10 ^ 5)
    TotalP44 = Application.SumIfs(PenSheet.Range("R1:R" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer4, _
    PenSheet.Range("N1:N" & LrPen), "BMT", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP45 = Application.SumIfs(PenSheet.Range("R1:R" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer4, _
    PenSheet.Range("N1:N" & LrPen), "R-T", PenSheet.Range("M1:M" & LrPen), "SD0045")
    TotalP46 = Application.SumIfs(PenSheet.Range("R1:R" & LrPen), PenSheet.Range("A1:A" & LrPen), Customer4, _
    PenSheet.Range("N1:N" & LrPen), "R-S", PenSheet.Range("M1:M" & LrPen), "SD0045")
    PiSheet.Cells(i4, "X").Value = (TotalP44 + TotalP45 + TotalP46)
    PiSheet.Cells(i4, "Y").Value = PiSheet.Cells(i4, "U").Value + PiSheet.Cells(i4, "W").Value
    PiSheet.Cells(i4, "Z").Value = PiSheet.Cells(i4, "V").Value + PiSheet.Cells(i4, "X").Value
    Next i4
 
'First Final formatting
    Dim LrPi1A As Long
    Dim LrPi2A As Long
    Dim LrPi3A As Long
    Dim LrPi4A As Long
    PiSheet.Columns.AutoFit
    PiSheet.Rows("1:2").Insert Shift:=xlDown
    LrPi1A = PiSheet.Cells(PiSheet.Rows.Count, "A").End(xlUp).Row
    PiSheet.Range("A1:B2").Merge
    PiSheet.Range("A1").Value = "All Data"
    PiSheet.Range("C1:E2").Merge
    PiSheet.Range("C1").Value = Date
    PiSheet.Range("A" & LrPi1A + 1 & ":B" & LrPi1A + 1).Merge
    PiSheet.Range("A" & LrPi1A + 1 & ":B" & LrPi1A + 1).Value = "Grand Total"
    PiSheet.Range("A3").Value = "Customer_code"
    PiSheet.Range("B3").Value = "Customer_Name"
    PiSheet.Range("C3").Value = "Billed_Amt"
    PiSheet.Range("D3").Value = "Pending_Amt"
    PiSheet.Range("E3").Value = "Total_Amt"
    With PiSheet.Range("A1:E3")
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    PiSheet.Range("C" & LrPi1A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("C4:C" & LrPi1A))
    PiSheet.Range("D" & LrPi1A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("D4:D" & LrPi1A))
    PiSheet.Range("E" & LrPi1A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("E4:E" & LrPi1A))
    With PiSheet.Range("A" & LrPi1A + 1 & ":E" & LrPi1A + 1)
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    End With
    With PiSheet.Range("A" & LrPi1A + 1)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("A4:B" & LrPi1A)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("C4:E" & LrPi1A + 1)
    .NumberFormat = "#,##0.00"
    End With
    With PiSheet.Range("A1:E" & LrPi1A + 1)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    End With
    With PiSheet.Range("A4:E" & LrPi1A)
        .Sort Key1:=PiSheet.Range("E4"), Order1:=xlDescending, Header:=xlNo
    End With
     
'Second Final formatting
    LrPi2A = PiSheet.Cells(PiSheet.Rows.Count, "G").End(xlUp).Row
    PiSheet.Range("G1:H2").Merge
    PiSheet.Range("G1").Value = "Moulded Fittings"
    PiSheet.Range("I1:K2").Merge
    PiSheet.Range("I1").Value = Date
    PiSheet.Range("G" & LrPi2A + 1 & ":H" & LrPi2A + 1).Merge
    PiSheet.Range("G" & LrPi2A + 1 & ":H" & LrPi2A + 1).Value = "Grand Total"
    PiSheet.Range("G3").Value = "Customer_code"
    PiSheet.Range("H3").Value = "Customer_Name"
    PiSheet.Range("I3").Value = "Billed_Amt"
    PiSheet.Range("J3").Value = "Pending_Amt"
    PiSheet.Range("K3").Value = "Total_Amt"
    With PiSheet.Range("G1:K3")
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    PiSheet.Range("I" & LrPi2A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("I4:I" & LrPi2A))
    PiSheet.Range("J" & LrPi2A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("J4:J" & LrPi2A))
    PiSheet.Range("K" & LrPi2A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("K4:K" & LrPi2A))
    With PiSheet.Range("G" & LrPi2A + 1 & ":K" & LrPi2A + 1)
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    End With
    With PiSheet.Range("G" & LrPi2A + 1)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
     With PiSheet.Range("G4:H" & LrPi2A)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("I4:K" & LrPi2A + 1)
    .NumberFormat = "#,##0.00"
    End With
    With PiSheet.Range("G1:K" & LrPi2A + 1)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    End With
    With PiSheet.Range("G4:K" & LrPi2A)
        .Sort Key1:=PiSheet.Range("K4"), Order1:=xlDescending, Header:=xlNo
    End With
 
'Third Final formatting
    LrPi3A = PiSheet.Cells(PiSheet.Rows.Count, "M").End(xlUp).Row
    PiSheet.Range("M1:N2").Merge
    PiSheet.Range("M1").Value = "CPVC"
    PiSheet.Range("O1:Q2").Merge
    PiSheet.Range("O1").Value = Date
    PiSheet.Range("M" & LrPi3A + 1 & ":N" & LrPi3A + 1).Merge
    PiSheet.Range("M" & LrPi3A + 1 & ":N" & LrPi3A + 1).Value = "Grand Total"
    PiSheet.Range("M3").Value = "Customer_code"
    PiSheet.Range("N3").Value = "Customer_Name"
    PiSheet.Range("O3").Value = "Billed_Amt"
    PiSheet.Range("P3").Value = "Pending_Amt"
    PiSheet.Range("Q3").Value = "Total_Amt"
    With PiSheet.Range("M1:Q3")
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    PiSheet.Range("O" & LrPi3A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("O4:O" & LrPi3A))
    PiSheet.Range("P" & LrPi3A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("P4:P" & LrPi3A))
    PiSheet.Range("Q" & LrPi3A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("Q4:Q" & LrPi3A))
    With PiSheet.Range("M" & LrPi3A + 1 & ":Q" & LrPi3A + 1)
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    End With
    With PiSheet.Range("M" & LrPi3A + 1)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("M4:N" & LrPi3A)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("O4:Q" & LrPi3A + 1)
    .NumberFormat = "#,##0.00"
    End With
    With PiSheet.Range("M1:Q" & LrPi3A + 1)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    End With
    With PiSheet.Range("M4:Q" & LrPi3A)
        .Sort Key1:=PiSheet.Range("Q4"), Order1:=xlDescending, Header:=xlNo
    End With
 
'Fourth Final formatting
    LrPi4A = PiSheet.Cells(PiSheet.Rows.Count, "S").End(xlUp).Row
    PiSheet.Range("S1:T1").Merge
    PiSheet.Range("S1").Value = "Tanks"
    PiSheet.Range("U1:Z1").Merge
    PiSheet.Range("U1").Value = Date
    PiSheet.Range("S" & LrPi4A + 1 & ":T" & LrPi4A + 1).Merge
    PiSheet.Range("S" & LrPi4A + 1 & ":T" & LrPi4A + 1).Value = "Grand Total"
    PiSheet.Range("S2:S3").Merge
    PiSheet.Range("S2").Value = "Customer_code"
    PiSheet.Range("T2:T3").Merge
    PiSheet.Range("T2").Value = "Customer_Name"
    PiSheet.Range("U2:V2").Merge
    PiSheet.Range("U2").Value = "Billing"
    PiSheet.Range("U3").Value = "Amt"
    PiSheet.Range("V3").Value = "Ltr"
    PiSheet.Range("W2:X2").Merge
    PiSheet.Range("W2").Value = "Pending"
    PiSheet.Range("W3").Value = "Amt"
    PiSheet.Range("X3").Value = "Ltr"
    PiSheet.Range("Y2:Z2").Merge
    PiSheet.Range("Y2").Value = "Total"
    PiSheet.Range("Y3").Value = "Amt"
    PiSheet.Range("Z3").Value = "Ltr"
    With PiSheet.Range("S1:Z3")
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    PiSheet.Range("U" & LrPi4A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("U4:U" & LrPi4A))
    PiSheet.Range("V" & LrPi4A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("V4:V" & LrPi4A))
    PiSheet.Range("W" & LrPi4A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("W4:W" & LrPi4A))
    PiSheet.Range("X" & LrPi4A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("X4:X" & LrPi4A))
    PiSheet.Range("Y" & LrPi4A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("Y4:Y" & LrPi4A))
    PiSheet.Range("Z" & LrPi4A + 1).Value = Application.WorksheetFunction.Sum(PiSheet.Range("Z4:Z" & LrPi4A))
    With PiSheet.Range("S" & LrPi4A + 1 & ":Z" & LrPi4A + 1)
    .Font.Size = 14
    .Font.Bold = True
    .Interior.Color = RGB(200, 220, 240)
    End With
    With PiSheet.Range("S" & LrPi4A + 1)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("S4:T" & LrPi4A)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    With PiSheet.Range("U4:Z" & LrPi4A + 1)
    .NumberFormat = "#,##0.00"
    End With
    With PiSheet.Range("S1:Z" & LrPi4A + 1)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    End With
    With PiSheet.Range("S4:Z" & LrPi4A)
        .Sort Key1:=PiSheet.Range("Y4"), Order1:=xlDescending, Header:=xlNo
    End With
    
    PiSheet.Columns.AutoFit
    PiSheet.Activate
    PiSheet.Range("A1").Select
    VBook.Close
    SourceBook.Close
    SaleBook.Close
  
  Dim FilePath3 As String
  Dim FilePath4 As String
  FilePath3 = "Billing_Pending_murali" & TodaysDate & ".XLSM"
  FilePath4 = FolderPath & FilePath3
  ThisWorkbook.SaveCopyAs FilePath4
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Rows.Delete
    Next ws

    
End Sub


