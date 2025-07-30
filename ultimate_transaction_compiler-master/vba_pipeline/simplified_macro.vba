Option Explicit

Sub CreateOutput1()
    Dim vfileA As Variant
    Dim vfileB As Variant
    Dim vfileC As Variant
    Dim vfileD As Variant
    Dim vfileE As Variant
    Dim vfileF As Variant
    Dim vfileG As Variant
    Dim vfileH As Variant
    Dim sht As Worksheet
    Dim strpathsave As String
    Dim strpath As String
    Dim wbRFM As Workbook
    Dim shtRFM As Worksheet
    Dim strfilename As String
    Dim wbA As Workbook
    Dim wbB As Workbook
    Dim wbC As Workbook
    Dim wbD As Workbook
    Dim wbE As Workbook
    Dim wbF As Workbook
    Dim wbG As Workbook
    Dim wbH As Workbook
    Dim shtA As Worksheet
    Dim shtb As Worksheet
    Dim shtC As Worksheet
    Dim shtD As Worksheet
    Dim shtE As Worksheet
    Dim shtF As Worksheet
    Dim shtG As Worksheet
    Dim shtH As Worksheet
    Dim c As Object
    Dim lastrowA As Long
    Dim lastrowB As Long
    Dim lastrowC As Long
    Dim lastrowD As Long
    Dim lastrowE As Long
    Dim lastrowF As Long
    Dim lastrowG As Long
    Dim lastrowh As Long
    Dim Lastrow As Long
    Dim x As Long
    Dim strformula As String
    Dim shtOF1 As Worksheet
    Dim wbOF1 As Workbook
    Dim vfileoutput As Variant
    Dim lastrowOF1 As Long
    Dim wbO As Workbook
    Dim shtO As Worksheet
    Dim shtcalc As Worksheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Set sht = Sheets("RFM Analyzer")

    vfileoutput = sht.Range("B6")
    vfileB = sht.Range("B9")
    vfileF = sht.Range("B12")
    strpathsave = sht.Range("B15")
    strfilename = sht.Range("B18")

    If Dir(vfileoutput) = vbNullString Then
        MsgBox "Please select a valid path to the Output File 1."
        Exit Sub
    End If
    If Dir(vfileB) = vbNullString Then
        MsgBox "Please select a valid path to the export file B."
        Exit Sub
    End If
    If Dir(vfileF) = vbNullString Then
        MsgBox "Please select a valid path to the export file F."
        Exit Sub
    End If

    Set wbRFM = Workbooks.Add
    Set shtRFM = wbRFM.Sheets(1)
    shtRFM.Name = "Output"

    Set wbOF1 = Workbooks.Open(vfileoutput)
    Set shtOF1 = wbOF1.Sheets(1)
    Set wbB = Workbooks.Open(vfileB)
    Set shtb = wbB.Sheets(1)
    Set wbF = Workbooks.Open(vfileF)
    Set shtF = wbF.Sheets(1)

    ' Get last rows
    Set c = shtb.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If Not c Is Nothing Then
        lastrowB = c.Row
    End If

    Set c = shtOF1.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If Not c Is Nothing Then
        lastrowOF1 = c.Row
    End If

    Set c = shtF.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If Not c Is Nothing Then
        lastrowF = c.Row
    End If

    ' Copy data from OF1 to new sheet
    shtOF1.Copy after:=wbRFM.Sheets(wbRFM.Sheets.Count)
    wbRFM.Sheets(wbRFM.Sheets.Count).Name = "OF1"
    wbOF1.Close False

    ' Copy data from B to new sheet
    shtb.Copy after:=wbRFM.Sheets(wbRFM.Sheets.Count)
    wbRFM.Sheets(wbRFM.Sheets.Count).Name = "B"
    wbB.Close False

    ' Copy data from F to new sheet
    shtF.Copy after:=wbRFM.Sheets(wbRFM.Sheets.Count)
    wbRFM.Sheets(wbRFM.Sheets.Count).Name = "F"
    wbF.Close False

    ' Add calc sheet
    wbRFM.Sheets.Add after:=wbRFM.Sheets(wbRFM.Sheets.Count)
    wbRFM.Sheets(wbRFM.Sheets.Count).Name = "Calc"

    ' Set up Output sheet headers
    shtRFM.Range("A1") = "Constituent ID"
    shtRFM.Range("B1") = "Contact Channel Status"
    shtRFM.Range("C1") = "Current employer Name"
    shtRFM.Range("D1") = "Most Recent DS Score in Database"
    shtRFM.Range("E1") = "Most Recent DS Weatlh Based Capacity in Database"
    shtRFM.Range("F1") = "Current Portfolio Assignment in Database"
    shtRFM.Range("G1") = "Age Range"
    shtRFM.Range("H1") = "Race"
    shtRFM.Range("I1") = "Gender"
    shtRFM.Range("J1") = "MSA"
    shtRFM.Range("K1") = "RFM Score"
    shtRFM.Range("L1") = "Recency Score"
    shtRFM.Range("M1") = "Frequency Score"
    shtRFM.Range("N1") = "Monetary Score"
    shtRFM.Range("O1") = "Recency Criteria"
    shtRFM.Range("P1") = "Frequency Criteria"
    shtRFM.Range("Q1") = "Monetary Criteria"
    shtRFM.Range("R1") = "RFM Percentile"
    shtRFM.Range("S1") = "Recency Percentile"
    shtRFM.Range("T1") = "Frequency Percentile"
    shtRFM.Range("U1") = "Monetary Percentile"
    shtRFM.Range("V1") = "Total Number of Gifts"
    shtRFM.Range("W1") = "Giving Segment A"
    shtRFM.Range("X1") = "Giving Segment B"
    shtRFM.Range("Y1") = "Primary Giving Platform"
    shtRFM.Range("Z1") = "Primary Giving Platform %"
    shtRFM.Range("AA1") = "Lifetime Giving Sum"
    shtRFM.Range("AB1") = "Lifetime Giving Range A"
    shtRFM.Range("AC1") = "Lifetime Giving Range B"
    shtRFM.Range("AD1") = "Digital Monthly Indicator"
    shtRFM.Range("AE1") = "Last Monthly Gift Date"
    shtRFM.Range("AF1") = "Last Monthly Gift Date Range"
    shtRFM.Range("AG1") = "Last Monthly Gift Amount"
    shtRFM.Range("AH1") = "Last Monthly Giving Amount Range"
    shtRFM.Range("AI1") = "Last Gift Date"
    shtRFM.Range("AJ1") = "Last Gift Date Range"
    shtRFM.Range("AK1") = "Last Gift Amount"
    shtRFM.Range("AL1") = "Last Gift Amount Range"
    shtRFM.Range("AM1") = "Last Gift Giving Platform"
    shtRFM.Range("AN1") = "Last Gift Campaign Name"
    shtRFM.Range("AO1") = "Last Gift Appeal Name"
    shtRFM.Range("AP1") = "Largest Gift Date"
    shtRFM.Range("AQ1") = "Largest Gift Date Range"
    shtRFM.Range("AR1") = "Largest Gift Amount"
    shtRFM.Range("AS1") = "Largest Gift Amount Range"
    shtRFM.Range("AT1") = "Largest Gift Platform"
    shtRFM.Range("AU1") = "Largest Gift Campaign Name"
    shtRFM.Range("AV1") = "Largest Gift Appeal Name"
    shtRFM.Range("AW1") = "First Gift Date"
    shtRFM.Range("AX1") = "First Gift Date Range A"
    shtRFM.Range("AY1") = "First Gift Date Range B"
    shtRFM.Range("AZ1") = "First Gift Amount"
    shtRFM.Range("BA1") = "First Gift Amount Range"
    shtRFM.Range("BB1") = "First Gift Platform"
    shtRFM.Range("BC1") = "First Gift Campaign Name"
    shtRFM.Range("BD1") = "First Gift Appeal Name"
    shtRFM.Range("BE1") = "Longitude"
    shtRFM.Range("BF1") = "Latitude"
    shtRFM.Range("BG1") = "First Gift Digital Monthly"
    shtRFM.Range("BH1") = "Largest Gift Digital Monthly"
    shtRFM.Range("BI1") = "Last Gift Digital Monthly"

    ' Copy data from OF1 to Output
    wbRFM.Sheets("OF1").Range("A2:F" & lastrowOF1).Copy shtRFM.Range("A2")
    wbRFM.Sheets("OF1").Range("H2:K" & lastrowOF1).Copy shtRFM.Range("G2")
    wbRFM.Sheets("OF1").Range("L2:M" & lastrowOF1).Copy shtRFM.Range("BE2:BF2")

    ' Set up formulas
    shtRFM.Range("K2") = "=sum(L2:N2)"
    shtRFM.Range("L2") = "=IFERROR(RANK.EQ(O2,INDEX(A:Q,0,COLUMN(O2)),1)/COUNT(INDEX(A:Q,0,COLUMN(O2)))*10,0)"
    shtRFM.Range("M2") = "=IFERROR(RANK.EQ(P2,INDEX(A:Q,0,COLUMN(P2)),1)/COUNT(INDEX(A:Q,0,COLUMN(P2)))*10,0)"
    shtRFM.Range("N2") = "=IFERROR(RANK.EQ(Q2,INDEX(A:Q,0,COLUMN(Q2)),1)/COUNT(INDEX(A:Q,0,COLUMN(Q2)))*10,0)"
    shtRFM.Range("O2") = "=iferror(MAX(FILTER(B!D:D,B!A:A=$A2)),""None"")"
    shtRFM.Range("P2") = "=IFERROR(COUNT(UNIQUE(FILTER(B!B:B,B!A:A=$A2))),0)"
    shtRFM.Range("Q2") = "=SUMIF(B!A:A,""="" & $A2,B!E:E)"

    ' Copy formulas down
    shtRFM.Range("K2:N2").NumberFormat = "##.00"
    shtRFM.Range("K2:BD2").Copy shtRFM.Range("K3:BD" & lastrowOF1)
    shtRFM.Range("BG2:BI2").Copy shtRFM.Range("BG3:BI" & lastrowOF1)

    ' Set number formats
    shtRFM.Range("O:O").NumberFormat = "mm/dd/yyyy"
    shtRFM.Range("Q:Q").NumberFormat = "$#,###,###,###,###.00"
    shtRFM.Range("Y:Y").NumberFormat = "#.00%"
    shtRFM.Range("Z:Z").NumberFormat = "#.00%"
    shtRFM.Range("AA:AA").NumberFormat = "$#,###,###,###,###.00"
    shtRFM.Range("AE:AE").NumberFormat = "mm/dd/yyyy"
    shtRFM.Range("AI:AI").NumberFormat = "mm/dd/yyyy"
    shtRFM.Range("AK:AK").NumberFormat = "$#,###,###,###,###.00"
    shtRFM.Range("AP:AP").NumberFormat = "mm/dd/yyyy"
    shtRFM.Range("AR:AR").NumberFormat = "$#,###,###,###,###.00"
    shtRFM.Range("AW:AW").NumberFormat = "mm/dd/yyyy"

    ' Delete temporary sheets
    wbRFM.Sheets("OF1").Delete
    wbRFM.Sheets("B").Delete
    wbRFM.Sheets("F").Delete
    wbRFM.Sheets("Calc").Delete

    ' Save the workbook
    wbRFM.SaveAs strpathsave & Application.PathSeparator & strfilename & ".xlsx"

    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done!"

End Sub
