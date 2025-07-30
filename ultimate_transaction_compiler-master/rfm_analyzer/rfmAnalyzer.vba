 Option Explicit
Global bError As Boolean
Global blShowProgress As Boolean
Sub PathtoExportFiles()
Dim x As Long
Dim fdlg As Object
Dim vfolder As Variant
Dim sht As Worksheet

Set sht = Sheets("RFM Analyzer")

#If Mac Then
    vfolder = Select_Folder_On_Mac("B", 9)
#Else
    Set fdlg = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlg
        .AllowMultiSelect = False
        .Title = "Select Path to Export Files..."
        If .Show <> 0 Then
            vfolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
#End If
sht.Range("B6") = vfolder
End Sub
Sub GetFileB()
#If Mac Then
    Call Select_File_Or_Files_Mac("B", 9)
#Else
    Call BrowseFile("B", 9)
#End If
End Sub
Sub GetFileF()
#If Mac Then
    Call Select_File_Or_Files_Mac("F", 12)
#Else
    Call BrowseFile("F", 12)
#End If
End Sub
Sub GetFileOF1()
#If Mac Then
    Call Select_File_Or_Files_Mac("1", 6)
#Else
    Call BrowseFile("1", 6)
#End If
End Sub


Sub PathToSave()
Dim x As Long
Dim fdlg As Object
Dim vfolder As Variant
Dim sht As Worksheet

Set sht = Sheets("RFM Analyzer")
#If Mac Then
    vfolder = Select_Folder_On_Mac()
#Else

    Set fdlg = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlg
        .AllowMultiSelect = False
        .Title = "Select Path to Save Files..."
        If .Show <> 0 Then
            vfolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
#End If
sht.Range("B15") = vfolder

End Sub
Function Select_Folder_On_Mac() As String
    Dim FolderPath As String
    Dim RootFolder As String
    Dim Scriptstr As String

    On Error Resume Next
    
    'Enter the Start Folder, Desktop in this example,
    'Use the second line to enter your own path
    RootFolder = MacScript("return POSIX path of (path to desktop folder) as String")
    
    'RootFolder = "/Users/rondebruin/Desktop/TestFolder/"
    
    'Make the path Colon seperated for using in MacScript
    RootFolder = MacScript("return POSIX file (""" & RootFolder & """) as string")
    'Make the Script string
        Scriptstr = "return POSIX path of (choose folder with prompt ""Select the folder""" & _
            " default location alias """ & RootFolder & """) as string"
    
    'Run the Script
    FolderPath = MacScript(Scriptstr)
    On Error GoTo 0
    
    Select_Folder_On_Mac = FolderPath
End Function

Sub BrowseFile(str As String, irow As Integer)
Dim x As Long
Dim fdlg As Object
Dim vfile As Variant
Dim sht As Worksheet

Set sht = Sheets("RFM Analyzer")
Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
With fdlg
    .AllowMultiSelect = False
    .Title = "Select Path to Export File " & str & "..."
    If .Show <> 0 Then
        vfile = .SelectedItems(1)
    Else
        Exit Sub
    End If
End With
sht.Range("B" & irow) = vfile

End Sub
Function Select_File_Or_Files_Mac(str As String, irow As Integer) As String
    'Select files in Mac Excel with the format that you want
    'Working in Mac Excel 2011 and 2016 and higher
    'Ron de Bruin, 20 March 2016
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim MySplit As Variant
    Dim N As Long
    Dim Fname As String
    Dim mybook As Workbook
    Dim OneFile As Boolean
    Dim FileFormat As String
    Dim sht As Worksheet
    
    Set sht = Sheets("RFM Analyzer")
    'In this example you can only select xlsx files
    'See my webpage how to use other and more formats.
     '   FileFormat = "{""public.comma-separated-values-text""}"
    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet"",""com.microsoft.Excel.xls""}"
   

    ' Set to True if you only want to be able to select one file
    ' And to False to be able to select one or more files
    OneFile = True

    On Error Resume Next
    MyPath = MacScript("return (path to desktop folder) as String")
    'Or use A full path with as separator the :
    'MyPath = "HarddriveName:Users::Desktop:YourFolder:"

    'Building the applescript string, do not change this
    If Val(Application.Version) < 15 Then
        'This is Mac Excel 2011
        If OneFile = True Then
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return theFile"
        Else
            MyScript = _
                "set applescript's text item delimiters to {ASCII character 10} " & vbNewLine & _
                "set theFiles to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                MyPath & """ with multiple selections allowed) as string" & vbNewLine & _
                "set applescript's text item delimiters to """" " & vbNewLine & _
                "return theFiles"
        End If
    Else
        'This is Mac Excel 2016 or higher
        If OneFile = True Then
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return posix path of theFile"
        Else
            MyScript = _
                "set theFiles to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                MyPath & """ with multiple selections allowed)" & vbNewLine & _
                "set thePOSIXFiles to {}" & vbNewLine & _
                "repeat with aFile in theFiles" & vbNewLine & _
                "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
                "end repeat" & vbNewLine & _
                "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
                "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
                "set text item delimiters to TID" & vbNewLine & _
                "return thePOSIXFiles"
        End If
    End If

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    'If you select one or more files MyFiles is not empty
    'We can do things with the file paths now like I show you below
    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        'MySplit = Split(MyFiles, Chr(10))
       Select_File_Or_Files_Mac = MyFiles
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
    End If
    sht.Range("B" & irow) = MyFiles
End Function
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
'Application.ScreenUpdating = False
frmprogress.Show vbModeless
blShowProgress = True
DoEvents
Application.Calculation = xlCalculationManual
Updateprogress , 0.5, "Copying and calculating Data...", True

Set sht = Sheets("RFM Analyzer")

vfileoutput = sht.Range("B6")
vfileB = sht.Range("B9")
vfileF = sht.Range("B12")

strpathsave = sht.Range("B15")
strfilename = sht.Range("B18")

If vfileoutput = vbNullString Then
    MsgBox "Please select a path to Output File 1."
    Call BrowseFile("1", 6)
    Exit Sub
End If
If vfileB = vbNullString Then
    MsgBox "Please select a path to Export File B."
    Call BrowseFile("B", 9)
    Exit Sub
End If
If vfileF = vbNullString Then
    MsgBox "Please select a path to Export File F."
    Call BrowseFile("F", 12)
    Exit Sub
End If




If strpathsave = vbNullString Then
    MsgBox "Please select a save as path for the output file(s)."
    Call PathToSave
    Exit Sub
End If

If Dir(vfileoutput) = vbNullString Then
    MsgBox "Please select a valid path to the Output File 1."
    Call BrowseFile("1", 6)
    Exit Sub
End If
If Dir(vfileB) = vbNullString Then
    MsgBox "Please select a valid path to the export file B."
    Call BrowseFile("B", 9)
    Exit Sub
End If
If Dir(vfileF) = vbNullString Then
    MsgBox "Please select a valid path to the export file F."
    Call BrowseFile("F", 12)
    Exit Sub
End If

If Dir(strpathsave, vbDirectory) = vbNullString Then
    MsgBox "Please select a valid save as path for the output file(s)."
    Call PathToSave
    Exit Sub
End If

If strfilename = vbNullString Then
    MsgBox "Please enter a save as file name in B15."
    sht.Range("B15").Select
    Exit Sub
End If


Set wbRFM = Workbooks.Add
Set shtRFM = wbRFM.Sheets(1)
shtRFM.Name = "Output"

Set wbOF1 = Workbooks.Open(vfileoutput)
Set shtOF1 = wbOF1.Sheets(1)
'Set wbB = Workbooks.Open(strpath & Application.PathSeparator & vfileB)
'Set shtB = wbB.Sheets(1)
Set wbB = Workbooks.Open(vfileB)
Set shtb = wbB.Sheets(1)
Set wbF = Workbooks.Open(vfileF)
Set shtF = wbF.Sheets(1)



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


DoEvents
Updateprogress , x / lastrowOF1, "Copying Data...", True


shtOF1.Copy after:=wbRFM.Sheets(wbRFM.Sheets.Count)
wbRFM.Sheets(wbRFM.Sheets.Count).Name = "OF1"
wbOF1.Close False

shtb.Copy after:=wbRFM.Sheets(wbRFM.Sheets.Count)
wbRFM.Sheets(wbRFM.Sheets.Count).Name = "B"
wbB.Close False

Set shtb = wbRFM.Sheets("B")
shtb.Activate
shtb.Range("I2") = "=COUNTIFS(C:C,""="" & $C2,A:A,""="" & $A2)/COUNTIF(A:A,""="" & $A2)"
shtb.Range("I2").NumberFormat = "#.00%"
shtb.Range("I2").Copy shtb.Range("I3:I" & lastrowB)

shtF.Copy after:=wbRFM.Sheets(wbRFM.Sheets.Count)
wbRFM.Sheets(wbRFM.Sheets.Count).Name = "F"
wbF.Close False

wbRFM.Sheets.Add after:=wbRFM.Sheets(wbRFM.Sheets.Count)
wbRFM.Sheets(wbRFM.Sheets.Count).Name = "Calc"

Set shtcalc = wbRFM.Sheets("calc")


shtRFM.Range("A1") = "Constituent ID"
shtRFM.Range("B1") = "Contact Channel Status"
shtRFM.Range("C1") = "Current employer Name"
shtRFM.Range("D1") = "Most Recent DS Score in Database"
shtRFM.Range("E1") = "Most Recent DS Weatlh Based Capacity in Database"
shtRFM.Range("F1") = "Current Portfolio Assignment in Database"
'shtRFM.Range("G1") = "Age"
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

wbRFM.Sheets("OF1").Range("A2:F" & lastrowOF1).Copy shtRFM.Range("A2")
wbRFM.Sheets("OF1").Range("H2:K" & lastrowOF1).Copy shtRFM.Range("G2")
wbRFM.Sheets("OF1").Range("L2:M" & lastrowOF1).Copy shtRFM.Range("BE2:BF2")

'shtRFM.Range("O2") = "=INDEX(B!$A:$D, MATCH(MAX(B!$D:$D), B!$D:$D,0),4)"

shtRFM.Range("K2") = "=sum(L2:N2)"
shtRFM.Range("L2") = "=IFERROR(RANK.EQ(O2,INDEX(A:Q,0,COLUMN(O2)),1)/COUNT(INDEX(A:Q,0,COLUMN(O2)))*10,0)"
shtRFM.Range("M2") = "=IFERROR(RANK.EQ(P2,INDEX(A:Q,0,COLUMN(P2)),1)/COUNT(INDEX(A:Q,0,COLUMN(P2)))*10,0)"
shtRFM.Range("N2") = "=IFERROR(RANK.EQ(Q2,INDEX(A:Q,0,COLUMN(Q2)),1)/COUNT(INDEX(A:Q,0,COLUMN(Q2)))*10,0)"

shtRFM.Range("O2") = "=iferror(MAX(FILTER(B!D:D,B!A:A=$A2)),""None"")"
shtRFM.Range("P2") = "=IFERROR(COUNT(UNIQUE(FILTER(B!B:B,B!A:A=$A2))),0)"
shtRFM.Range("Q2") = "=SUMIF(B!A:A,""="" & $A2,B!E:E)"
shtRFM.Range("R2") = "=iferror(IF(K2>=PERCENTILE.INC(K:K,99%),""A. Top 1%"",IF(K2>=PERCENTILE.INC(K:K,98%),""B. Top 2%"",IF(K2>=PERCENTILE.INC(K:K,90%),""C. Top 10%"",IF(K2>=PERCENTILE.INC(K:K,50%),""D. Top 50%"",""E. Bottom 50%"")))),""None"")"
shtRFM.Range("S2") = "=iferror(IF(L2>=PERCENTILE.INC(L:L,99%),""A. Top 1%"",IF(L2>=PERCENTILE.INC(L:L,98%),""B. Top 2%"",IF(L2>=PERCENTILE.INC(L:L,90%),""C. Top 10%"",IF(L2>=PERCENTILE.INC(L:L,50%),""D. Top 50%"",""E. Bottom 50%"")))),""None"")"
shtRFM.Range("T2") = "=iferror(IF(M2>=PERCENTILE.INC(M:M,99%),""A. Top 1%"",IF(M2>=PERCENTILE.INC(M:M,98%),""B. Top 2%"",IF(M2>=PERCENTILE.INC(M:M,90%),""C. Top 10%"",IF(M2>=PERCENTILE.INC(M:M,50%),""D. Top 50%"",""E. Bottom 50%"")))),""None"")"
shtRFM.Range("U2") = "=iferror(IF(N2>=PERCENTILE.INC(N:N,99%),""A. Top 1%"",IF(N2>=PERCENTILE.INC(N:N,98%),""B. Top 2%"",IF(N2>=PERCENTILE.INC(N:N,90%),""C. Top 10%"",IF(N2>=PERCENTILE.INC(N:N,50%),""D. Top 50%"",""E. Bottom 50%"")))),""None"")"

shtRFM.Range("V2") = "=P2"
shtRFM.Range("W2") = "=iferror(IF(V2=0,""A. No Gifts"", IF(V2= 1, ""A. First Gift Segment"", IF(V2= 2, ""B. Second Gift Segment"", IF(v2= 3, ""C. Third Gift Segment"", IF(V2= 4, ""D. Fourth Gift Segment"", IF(V2= 5, ""E. Ffifth Gift Segment"", IF(V2=6, ""F. Sixth Gift Segment"", IF(V2= 7, ""G. Seventh Gift Segment"", IF(V2> 7 , ""H. Seventh Gift Segment +""))))))))),""None"")"
shtRFM.Range("X2") = "=iferror(if(V2=0,""None"",IF(V2<13, ""A. Less than 13"", IF(AND(V2>=13,V2<=17), ""A. 13-17 gifts"", IF(and(v2>=18,V2<=22), ""B. 18-22 gifts"", IF(And(v2>=23,V2<=27),""C. 23-27 gifts"", IF(And(V2>=28,V2<=32),""D. 28-32 gifts"", IF(and(v2>=33,v2<=36), ""F. 33-36 gifts"", If(V2>=37, ""G. 37+ gifts"")))))))),""None"")"



Updateprogress , 0.5, "Processing Primary Giving Platforms...", True
'

shtRFM.Range("Y2").Formula2 = "=IFERROR(INDEX(B!C:C,MATCH(1,(A2=B!A:A)*(MAXIFS(B!I:I,B!A:A,""="" & $A2)=B!I:I),0)),""None"")"
shtRFM.Range("Z2") = "=iferror(COUNTIFS(B!C:C,""="" & Y2,B!A:A,""="" & A2)/COUNTIF(B!A:A,""="" & A2),""None"")"
shtRFM.Range("AA2") = "=Q2"
shtRFM.Range("AB2") = "=iferror(if(Q2=0,""None"",if(Q2<10,""A. <$10"", if(Q2<100,""B. <$100"",if(Q2<1000,""C. <$1k"", if(Q2<10000,""D. $<10k"", if(Q2<100000,""E. $<100k"", if(q2<1000000,""F. <$1M"",if(Q2<10000000,""G. <$10M"",if(Q2<100000000,""H. <$100M""))))))))),""None"")"
strformula = "=iferror(if(AA2=0,""None"",IF(AA2<=4.99,""A. <$5"",IF(AND(AA2>4.99,AA2<=9.99),""B. $5 - $9.99"",IF(AND(AA2>9.99,AA2<=24.99),""C. $10 - $24"",IF(AND(AA2>24.99,AA2<=49.99),""D. $25 - $49"",IF(AND(AA2>49.99,AA2<=99.99),""E. $50 - $99"",IF(AND(AA2>99.99,AA2<=249.99),""F. $100 - $249"",IF(AND(AA2>249.99,AA2<=499.99),""G. $250 - $499"",IF(AND(AA2>499.99,AA2<=999.99),""H. $500 - $999"",IF(AND(AA2>999.99,AA2<=2499.99),""I. $1,000 - $2,499""," & _
"IF(AND(AA2>2499.99,AA2<=4999.99),""J. $2,500 - $4,999"",IF(AND(AA2>4999.99,AA2<=9999.99),""K. $5,000 - $9,999"",IF(AND(AA2>9999.99,AA2<=24999.99),""L. $10,000 - $24,999"",IF(AND(AA2>24999.99,AA2<=49999.99),""M. $25,000 - $49,999"",IF(AND(AA2>49999,AA2<=99999),""N. $50,000-$99,999"", IF(AND(AA2>99999,AA2<=249999),""O. $100,00-$249,999"",IF(AND(AA2>249999,AA2<=499999),""P. $250,000 - $499,999"",IF(AND(AA2>499999,AA2<=999999),""Q. $500,000-$999,999"",IF(AND(AA2>999999,AA2<=2499999),""R. $1,000,000 - $2499999"",IF(AND(AA2>2499999,AA2<=4999999),""S. $2,500,000 - $4,999,999"",IF(AND(AA2>4999999,AA2<=9999999),""T. $5,000,000 - $9,999,999"", IF(AND(AA2" & _
">9999999,AA2<24999999),""U. $10,000,000 - $24,999,999"", IF(AND(AA2>24999999,AA2<=49999999),""V. $25,000,000 - $49,999,999"",IF(AND(AA2>49999999,AA2<=99999999),""W. $50,00,000 - $99,999,999"",IF(AND(AA2>99999999,AA2<=249999999),""X. $100,000,000 - $249,999,999"",IF(AND(AA2>249999999,AA2<=499999999),""Y. $250,000,000 - $499,999,999"", IF(AND(AA2>499999999,AA2<=999999999),""Z. $500,000,000 - $999,999,999"", ""AA. $1,000,000+""))))))))))))))))))))))))))),""None"")"
shtRFM.Range("AC2") = strformula

shtRFM.Range("AD2") = "=IF(iserror(Vlookup(A2,F!A:A,1,false)),""Not Digital Monthly"",""Digital Monthly"")"
shtRFM.Range("AE2") = "=IFERROR(IF(MAXIFS(F!F:F,F!A:A,""="" & $A2)=0,""No Digital Monthly Giving History"",MAXIFS(F!F:F,F!A:A,""="" & $A2)),""No Digital Monthly Giving History"")"
shtRFM.Range("AF2").Formula2 = "=IFERROR(if(not(isNumber(AE2)),""L. None"", if(and(AE2=today(),AE2<now()),""A. Earlier Today"",if(AE2=today()-1,""B. Yesterday"",if(and(AE2>=TODAY() - (WEEKDAY(TODAY()) - 1), AE2<=TODAY() - (WEEKDAY(TODAY()) - 1)+6), ""C. This Week"",if(and(AE2>=TODAY() - (WEEKDAY(TODAY()) - 1)-7, AE2<=TODAY() - (WEEKDAY(TODAY()) - 1)-1), ""D. Last Week"",if(and(AE2>=eomonth(today(),-1)+1, AE2<=eomonth(today(),0)), ""E. This Month"",if(and(AE2>=eomonth(today(),-2)+1, AE2<=eomonth(today(),-1)), ""F. Last Month""," & _
"if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),0)>=AE2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)<AE2),""G. This Quarter"",if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)>=AE2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-6)<AE2),""H. Last Quarter"",if(year(AE2)=year(today()),""I. This Year"",if(year(AE2)=year(today())-1,""J. Last Year"",if(year(AE2)<year(today())-1,""K. Before Last Year"",""L. None"")))))))))))),""None"")"

shtRFM.Range("AG2") = "=iferror(vlookup(A2,F!A:H,7,false),""No Digital Monthly Giving History"")"
shtRFM.Range("AH2") = "=iferror(if(AG2=""No Digital Monthly Giving History"",""No Digital Monthly Giving History"",IF(AG2<=4.99,""A. <$5"",IF(AND(AG2>4.99,AG2<=9.99),""B. $5 - $9.99"",IF(AND(AG2>9.99,AG2<=24.99),""C. $10 - $24"",IF(AND(AG2>24.99,AG2<=49.99),""D. $25 - $49"",IF(AND(AG2>49.99,AG2<=99.99),""E. $50 - $99"",IF(AND(AG2>99.99,AG2<=249.99),""F. $100 - $249"",IF(AND(AG2>249.99,AG2<=499.99),""G. $250 - $499"",IF(AND(AG2>499.99,AG2<=999.99),""H. $500 - $999"",IF(AND(AG2>999.99,AG2<=2499.99),""I. $1,000 - $2,499"",IF(AND(AG2>2499.99,AG2<=4999.99),""J. $2,500 - $4,999""," & _
"IF(AND(AG2>4999.99,AG2<=9999.99),""K. $5,000 - $9,999"",IF(AND(AG2>9999.99,AG2<=24999.99),""L. $10,000 - $24,999"",IF(AND(AG2>24999.99,AG2<=49999.99),""M. $25,000 - $49,999"",IF(AND(AG2>49999,AG2<=99999),""N. $50,000- $99,999"",IF(AND(AG2>99999,AG2<=249999),""O. $100,00-$249,999"",IF(AND(AG2>249999,AG2<=499999),""P. $250,000 - $499,999"",IF(AND(AG2>499999,AG2<=999999),""Q. $500,000-$999,999"",IF(AND(AG2>999999,AG2<=2499999),""R. $1,000,000 - $2,499,999"",IF(AND(AG2>2499999,AG2<=4999999),""S. $2,500,000 - $4,999,999"",IF(AND(AG2>4999999,AG2<=9999999),""T. $5,000,000 - $9,999,999""," & _
"IF(AND(AG2>9999999,AG2<=24999999),""U. $10,000,000 - $24,999,999"",IF(AND(AG2>24999999,AG2<=49999999),""V. $25,000,000 - $49,999,999"",IF(AND(AG2>49999999,AG2<=99999999),""W. $50,00,000 - $99,999,999"",IF(AND(AG2>99999999,AG2<=249999999),""X. $100,000,000 - $249,999,999"",IF(AND(AG2>249999999,AG2<=499999999),""Y. $250,000,000 - $499,999,999"",IF(AND(AG2>499999999,AG2<=999999999),""Z. 500,000,000 - $999,999,999"",""AA. $1,000,000+""))))))))))))))))))))))))))),""None"")"
shtRFM.Range("AI2") = "=IFERROR(IF(MAXIFS(B!D:D,B!A:A,""="" & A2)=0,""None"",MAXIFS(B!D:D,B!A:A,""="" & A2)),""None"")"

shtRFM.Range("AJ2").Formula2 = "=IFERROR(if(not(isNumber(AI2)),""L. None"", if(and(AI2=today(),AI2<now()),""A. Earlier Today"",if(AI2=today()-1,""B. Yesterday"",if(and(AI2>=TODAY() - (WEEKDAY(TODAY()) - 1), AI2<=TODAY() - (WEEKDAY(TODAY()) - 1)+6), ""C. This Week"",if(and(AI2>=TODAY() - (WEEKDAY(TODAY()) - 1)-7, AI2<=TODAY() - (WEEKDAY(TODAY()) - 1)-1), ""D. Last Week"",if(and(AI2>=eomonth(today(),-1)+1, AI2<=eomonth(today(),0)), ""E. This Month"",if(and(AI2>=eomonth(today(),-2)+1, AI2<=eomonth(today(),-1)), ""F. Last Month""," & _
"if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),0)>=AI2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)<AI2),""G. This Quarter"",if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)>=AI2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-6)<AI2),""H. Last Quarter"",if(year(AI2)=year(today()),""I. This Year"",if(year(AI2)=year(today())-1,""J. Last Year"",if(year(AI2)<year(today())-1,""K. Before Last Year"",""L. None"")))))))))))),""None"")"

shtRFM.Range("AK2").Formula2 = "=iferror(INDEX(B!E:E,MATCH(1,(A2=B!A:A)*(AI2=B!D:D),0)),0)"
shtRFM.Range("AL2") = "=iferror(if(Ak2=0,""None"",IF(AK2<=4.99,""A. <$5"",IF(AND(AK2>4.99,AK2<=9.99),""B. $5 - $9.99"",IF(AND(AK2>9.99,AK2<=24.99),""C. $10 - $24"",IF(AND(AK2>24.99,AK2<=49.99),""D. $25 - $49"",IF(AND(AK2>49.99,AK2<=99.99),""E. $50 - $99"",IF(AND(AK2>99.99,AK2<=249.99),""F. $100 - $249"",IF(AND(AK2>249.99,AK2<=499.99),""G. $250 - $499"",IF(AND(AK2>499.99,AK2<=999.99),""H. $500 - $999"",IF(AND(AK2>999.99,AK2<=2499.99),""I. $1,000 - $2,499"",IF(AND(AK2>2499.99,AK2<=4999.99),""J. $2,500 - $4,999""," & _
"IF(AND(AK2>4999.99,AK2<=9999.99),""K. $5,000 - $9,999"",IF(AND(AK2>9999.99,AK2<=24999.99),""L. $10,000 - $24,999"",IF(AND(AK2>24999.99,AK2<=49999.99),""M. $25,000 - $49,999"",IF(AND(AK2>49999,AK2<=99999),""N. $50,000- $99,999"",IF(AND(AK2>99999,AK2<=249999),""O. $100,00-$249,999"",IF(AND(AK2>249999,AK2<=499999),""P. $250,000 - $499,999"",IF(AND(AK2>499999,AK2<=999999),""Q. $500,000-$999,999"",IF(AND(AK2>999999,AK2<=2499999),""R. $1,000,000 - $2,499,999"",IF(AND(AK2>2499999,AK2<=4999999),""S. $2,500,000 - $4,999,999"",IF(AND(AK2>4999999,AK2<=9999999),""T. $5,000,000 - $9,999,999,""," & _
"IF(AND(AK2>9999999,AK2<=24999999),""U. $10,000,000 - $24,999,999"",IF(AND(AK2>24999999,AK2<=49999999),""V. $25,000,000 - $49,999,999"",IF(AND(AK2>49999999,AK2<=99999999),""W. $50,00,000 - $99,999,999"",IF(AND(AK2>99999999,AK2<=249999999),""X. $100,000,000 - $249,999,999"",IF(AND(AK2>249999999,AK2<=499999999),""Y. $250,000,000 - $499,999,999"",IF(AND(AK2>499999999,AK2<=1000000000),""Z. $500,000,000 - $999,999,999"",""AA. $1,000,000,000+""))))))))))))))))))))))))))),""None"")"
shtRFM.Range("AM2").Formula2 = "=iferror(INDEX(B!C:C,MATCH(1,(A2=B!A:A)*(AI2=B!D:D),0)),""None"")"
shtRFM.Range("AN2").Formula2 = "=iferror(INDEX(B!F:F,MATCH(1,(A2=B!A:A)*(AI2=B!D:D),0)),""None"")"
shtRFM.Range("AO2").Formula2 = "=iferror(INDEX(B!G:G,MATCH(1,(A2=B!A:A)*(AI2=B!D:D),0)),""None"")"
'shtRFM.Range("AP2").Formula2 = "=MAXIFS(B!D:D,B!A:A,""="" & A2,B!E:E,""="" & AR2)"

shtRFM.Range("AP2").Formula2 = "=IFERROR(IF(MAXIFS(B!D:D,B!A:A,""="" & A2,B!E:E,""="" & AR2)=0,""None"",MAXIFS(B!D:D,B!A:A,""="" & A2,B!E:E,""="" & AR2)),""None"")"



 shtRFM.Range("AQ2").Formula2 = "=IFERROR(if(not(isNumber(AP2)),""L. None"", if(and(AP2=today(),AP2<now()),""A. Earlier Today"",if(AP2=today()-1,""B. Yesterday"",if(and(AP2>=TODAY() - (WEEKDAY(TODAY()) - 1), AP2<=TODAY() - (WEEKDAY(TODAY()) - 1)+6), ""C. This Week"",if(and(AP2>=TODAY() - (WEEKDAY(TODAY()) - 1)-7, AP2<=TODAY() - (WEEKDAY(TODAY()) - 1)-1), ""D. Last Week"",if(and(AP2>=eomonth(today(),-1)+1, AP2<=eomonth(today(),0)), ""E. This Month"",if(and(AP2>=eomonth(today(),-2)+1, AP2<=eomonth(today(),-1)), ""F. Last Month""," & _
"if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),0)>=AP2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)<AP2),""G. This Quarter"",if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)>=AP2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-6)<AP2),""H. Last Quarter"",if(year(AP2)=year(today()),""I. This Year"",if(year(AP2)=year(today())-1,""J. Last Year"",if(year(AP2)<year(today())-1,""K. Before Last Year"",""L. None"")))))))))))),""None"")"

shtRFM.Range("AR2").Formula2 = "=MAXIFS(B!E:E,B!A:A,""="" & A2)"
strformula = "=iferror(IF(AR2=0,""None"",IF(AR2<=4.99,""A. <$5"",IF(AND(AR2>4.99,AR2<=9.99),""B. $5 - $9.99"",IF(AND(AR2>9.99,AR2<=24.99),""C. $10 - $24"",IF(AND(AR2>24.99,AR2<=49.99),""D. $25 - $49"",IF(AND(AR2>49.99,AR2<=99.99),""E. $50 - $99"",IF(AND(AR2>99.99,AR2<=249.99),""F. $100 - $249"",IF(AND(AR2>249.99,AR2<=499.99),""G. $250 - $499""," & _
"IF(AND(AR2>499.99,AR2<=999.99),""H. $500 - $999"",IF(AND(AR2>999.99,AR2<=2499.99),""I. $1,000 - $2,499"",IF(AND(AR2>2499.99,AR2<=4999.99),""J. $2,500 - $4,999"",IF(AND(AR2>4999.99,AR2<=9999.99),""K. $5,000 - $9,999"",IF(AND(AR2>9999.99,AR2<=24999.99),""L. $10,000 - $24,999"",IF(AND(AR2>24999.99,AR2<=49999.99),""M. $25,000 - $49,999""," & _
"IF(AND(AR2>49,999,AR2<=99999),""N. $50,000- $99,999"",IF(AND(AR2>99999,AR2<=249999),""O. $100,00-$249,999"",IF(AND(AR2>249999,AR2<=499999), ""P. $250,000 - $499,999"",IF(AND(AR2>499999,AR2<=999999),""Q. $500,000-$999,999"",IF(AND(AR2>999999,AR2<=2499999),""R. $1,000,000 - $2,499,999"",IF(AND(AR2>2499999,AR2<=4999999),""S. $2,500,000 - $4,999,999"",IF(AND(AR2>4999999,AR2<=9999999), ""T. $5,000,000 - $9,999,999"",IF(AND(AR2>9999999,AR2<=24999999), ""U. $10,000,000 - $24,999,999"",IF(AND(AR2>24999999,AR2<=49999999), ""V. $25,000,000 - $49,999,999"",IF(AND(AR2>49999999,AR2<=99999999), ""W. $50,00,000 - $99,999,999"",IF(AND(AR2>99999999,AR2<=249999999),""X. $100,000,000 - $249,999,999"",IF(AND(AR2>249999999,AR2<=499999999),""Y. $250,000,000 - $499,999,999"",IF(AND(AR2>499999999,AR2<=999999999),""Z. $500,000,000 - $999,999,999"", ""AA. $1,000,000+""))))))))))))))))))))))))))),""None"")"
shtRFM.Range("AS2").Formula2 = strformula
shtRFM.Range("AT2").Formula2 = "=iferror(INDEX(B!C:C,MATCH(1,(A2=B!A:A)*(AP2=B!D:D),0)),""None"")"
shtRFM.Range("AU2").Formula2 = "=iferror(INDEX(B!F:F,MATCH(1,(A2=B!A:A)*(AP2=B!D:D),0)),""None"")"
shtRFM.Range("AV2").Formula2 = "=iferror(INDEX(B!G:G,MATCH(1,(A2=B!A:A)*(AP2=B!D:D),0)),""None"")"
'shtRFM.Range("AW2").Formula2 = "=MINIFS(B!D:D,B!A:A,""="" & A2)"
shtRFM.Range("AW2").Formula2 = "=IFERROR(IF(MINIFS(B!D:D,B!A:A,""="" & A2)=0,""None"",MINIFS(B!D:D,B!A:A,""="" & A2)),""None"")"

strformula = "=iferror(if(Aw2=0,""None"",IF(YEAR(AW2)<2017,""A. Before 2017."",IF(AND(YEAR(AW2)>=2017,YEAR(AW2)<2019),""B 2017-2019."",IF(YEAR(AW2)=2019,""C. 2019."",IF(AND(AW2>=datevalue(""1/1/2020""),AW2<=datevalue(""5/25/2020"")),""D. 2020 Pre Anguish and Action (January - May 25, 2020)."",IF(AND(AW2>=datevalue(""5/26/2020""),AW2<=datevalue(""5/31/2020"")),""E. May 2020 Anguish and Action (May 26 2020 - May 31 2020)."",IF(AND(AW2>=datevalue(""6/1/2020""),AW2<=datevalue(""6/30/2020"")),""F. June 2020 Anguish and Action (June 1 2020 - June 30 2020)."",IF(AND(AW2>=datevalue(""7/1/2020""),AW2<=datavalue(""7/31/20"")),""G. July 2020 Anguish and Action (July 1 2020 - July 31 2020)."",IF(AND(AW2>=datevalue(""8/1/2020""), AW2<=datevalue(""8/31/2020"")),""H. August 2020 Anguish And Action (August 1, 2020 - August 31, 2020)."",IF(AND(AW2>=datevalue(""9/1/2020""),AW2<=datevalue(""12/31/2020""))," & _
"""J. September - December 2020 (September 1 2020 - December 31, 2020."",IF(YEAR(AW2)=2021,""K. 2021."",IF(YEAR(AW2)=2022,""L. 2022."","""")))))))))))),""None"")"
shtRFM.Range("AX2").Formula2 = strformula

shtRFM.Range("AY2").Formula2 = "=IFERROR(if(not(isNumber(AW2)),""L. None"", if(and(AW2=today(),AW2<now()),""A. Earlier Today"",if(AW2=today()-1,""B. Yesterday"",if(and(AW2>=TODAY() - (WEEKDAY(TODAY()) - 1), AW2<=TODAY() - (WEEKDAY(TODAY()) - 1)+6), ""C. This Week"",if(and(AW2>=TODAY() - (WEEKDAY(TODAY()) - 1)-7, AW2<=TODAY() - (WEEKDAY(TODAY()) - 1)-1), ""D. Last Week"",if(and(AW2>=eomonth(today(),-1)+1, AW2<=eomonth(today(),0)), ""E. This Month"",if(and(AW2>=eomonth(today(),-2)+1, AW2<=eomonth(today(),-1)), ""F. Last Month""," & _
"if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),0)>=AW2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)<AW2),""G. This Quarter"",if(and(eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-3)>=AW2,eomonth(date(year(today()),ROUNDUP(month(today())/3,0)*3,1),-6)<AW2),""H. Last Quarter"",if(year(AW2)=year(today()),""I. This Year"",if(year(AW2)=year(today())-1,""J. Last Year"",if(year(AW2)<year(today())-1,""K. Before Last Year"",""L. None"")))))))))))),""None"")"

shtRFM.Range("AZ2").Formula2 = "=iferror(INDEX(B!E:E,MATCH(1,(A2=B!A:A)*(AI2=B!D:D),0)),""None"")"


shtRFM.Range("BA2").Formula2 = "=iferror(if(OR(AZ2=""None"",AZ2=0),""None"",IF(AZ2<=4.99,""A. <$5"",IF(AND(AZ2>4.99,AZ2<=9.99),""B. $5 - $9.99"",IF(AND(AZ2>9.99,AZ2<=24.99),""C. $10 - $24"",IF(AND(AZ2>24.99,AZ2<=49.99),""D. $25 - $49"",IF(AND(AZ2>49.99,AZ2<=99.99),""E. $50 - $99"",IF(AND(AZ2>99.99,AZ2<=249.99),""F. $100 - $249"",IF(AND(AZ2>249.99,AZ2<=499.99),""G. $250 - $499"",IF(AND(AZ2>499.99,AZ2<=999.99),""H. $500 - $999"",IF(AND(AZ2>999.99,AZ2<=2499.99),""I. $1,000 - $2,499"",IF(AND(AZ2>2499.99,AZ2<=4999.99),""J. $2,500 - $4,999"",IF(AND(AZ2>4999.99,AZ2<=9999.99),""K. $5,000 - $9,999""," & _
"IF(AND(AZ2>9999.99,AZ2<=24999.99),""L. $10,000 - $24,999"",IF(AND(AZ2>24999.99,AZ2<=49999.99),""M. $25,000 - $49,999"",IF(AND(AZ2>49999.99,AZ2<=99999),""N. $50,000- $99,999"",IF(AND(AZ2>99999.99,AZ2<=249999), ""O. $100,00-$249,999"",IF(AND(AZ2>249999,AZ2<=499999), ""P. $250,000 - $499,999"",IF(AND(AZ2>499999,AZ2<=999999), ""Q. $500,000-$999,999"",IF(AND(AZ2>999999,AZ2<=2499999), ""R. $1,000,000 - $2,499,999"",IF(AND(AZ2>2499999,AZ2<=4999999),""S. $2,500,000 - $4,999,999"",IF(AND(AZ2>4999999,AZ2<=9999999),""T. $5,000,000 - $9,999,999""," & _
"IF(AND(AZ2>9999999,AZ2<=24999999), ""U. $10,000,000 - $24,999,999"",IF(AND(AZ2>24999999,AZ2<=49999999),""V. $25,000,000 - $49,999,999"",IF(AND(AZ2>49999999,AZ2<=99999999),""W. $50,00,000 - $99,999,999"",IF(AND(AZ2>99999999,AZ2<=249999999),""X. $100,000,000 - $249,999,999"",IF(AND(AZ2>249999999,AZ2<=499999999),""Y. $250,000,000 - $499,999,999"",IF(AND(AZ2>499999999,AZ2<=999999999),""Z. $500,000,000 - $999,999,999"", ""AA. $1,000,000+""))))))))))))))))))))))))))),""None"")"



'shtRFM.Range("BA2").Formula2 = "=iferror(INDEX(B!G:G,MATCH(1,(A2=B!A:A)*(AW2=B!D:D),0)),"""")"
shtRFM.Range("BB2").Formula2 = "=iferror(INDEX(B!C:C,MATCH(1,(A2=B!A:A)*(AW2=B!D:D),0)),""None"")"
shtRFM.Range("BC2").Formula2 = "=iferror(INDEX(B!F:F,MATCH(1,(A2=B!A:A)*(AW2=B!D:D),0)),""None"")"
shtRFM.Range("BD2").Formula2 = "=iferror(INDEX(B!G:G,MATCH(1,(A2=B!A:A)*(AW2=B!D:D),0)),""None"")"


shtRFM.Range("BG2").Formula2 = "=IF(ISERROR(VLOOKUP(INDEX(B!H:H,MATCH(1,(A2=B!A:A)*(MINIFS(B!D:D,B!A:A,""="" & A2)=B!D:D),0)),F!C:C,1,FALSE)),""None"",""Yes"")"
shtRFM.Range("BH2").Formula2 = "=IF(ISERROR(VLOOKUP(INDEX(B!H:H,MATCH(1,(A2=B!A:A)*(MAXIFS(B!D:D,B!A:A,""="" & A2,B!E:E,""="" & AR2)=B!D:D),0)),F!C:C,1,FALSE)),""None"",""Yes"")"
shtRFM.Range("BI2").Formula2 = "=IF(ISERROR(VLOOKUP(INDEX(B!H:H,MATCH(1,(A2=B!A:A)*(MAXIFS(B!D:D,B!A:A,""="" & A2)=B!D:D),0)),F!C:C,1,FALSE)),""None"",""Yes"")"




shtb.Activate
shtb.Range("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
shtb.Range("B:B").Select
Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
shtb.Range("H:H").Select
Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
shtb.Range("A:A").Select

shtb.Range("D:D").NumberFormat = "mm/dd/yyyy"
shtb.Range("D:D").Select
Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
shtb.Range("E:E").NumberFormat = "$#,##0.00"
shtb.Range("E:E").Select
Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

wbRFM.Sheets("F").Activate
wbRFM.Sheets("F").Range("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
        
wbRFM.Sheets("F").Range("C:C").Select
Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

wbRFM.Sheets("F").Range("B:B").NumberFormat = "mm/dd/yyyy"
wbRFM.Sheets("F").Range("E:E").NumberFormat = "mm/dd/yyyy"
wbRFM.Sheets("F").Range("F:F").NumberFormat = "mm/dd/yyyy"

wbRFM.Sheets("F").Range("D:D").NumberFormat = "$#,##0.00"
wbRFM.Sheets("F").Range("G:G").NumberFormat = "$#,##0.00"
wbRFM.Sheets("F").Range("E:E").Select
Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
wbRFM.Sheets("F").Range("F:F").Select
Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

wbRFM.Sheets("F").Range("D:D").Select
Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
wbRFM.Sheets("F").Range("G:G").Select
Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True


shtRFM.Range("K2:N2").NumberFormat = "##.00"
shtRFM.Range("K2:BD2").Copy shtRFM.Range("K3:BD" & lastrowOF1)
shtRFM.Range("BG2:BI2").Copy shtRFM.Range("BG3:BI" & lastrowOF1)
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

Application.Calculation = xlCalculationAutomatic

'Call GetRFM


Application.DisplayAlerts = False

Set c = shtRFM.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
If Not c Is Nothing Then
    Lastrow = c.Row
End If


'
'For x = 2 To Lastrow
'    DoEvents
'    Updateprogress , x / Lastrow, "Processing Primary Giving Platforms...", True
'    shtRFM.Range("Y" & x) = GetPrimaryGivingPlatform2(shtRFM.Range("A" & x), wbRFM)
'Next x

shtRFM.Activate
shtRFM.Range("AG:AG").Copy
shtRFM.Range("AG:AG").PasteSpecial xlPasteValues
shtRFM.Range("AG:AG").NumberFormat = "$#,###,###,###,###.00"

shtRFM.Range("AG:AG").Select
Selection.TextToColumns Destination:=Range("AG1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'

DoEvents
Updateprogress , 0.5, "Finalizing Formulas...", True
    
shtRFM.Activate
shtRFM.Range("A3:BZ" & lastrowOF1).Select
shtRFM.Activate
shtRFM.Range("A2:BZ" & lastrowOF1).Copy
shtRFM.Range("A2:BZ" & lastrowOF1).PasteSpecial xlPasteValues


wbRFM.SaveAs strpathsave & Application.PathSeparator & strfilename & ".xlsx"
shtRFM.Activate

DoEvents
Updateprogress , 0.95, "Just About Finished...", True



wbRFM.Sheets("OF1").Delete
wbRFM.Sheets("B").Delete
wbRFM.Sheets("F").Delete
Application.DisplayAlerts = True

Set shtOF1 = Nothing
Set shtb = Nothing
Set shtF = Nothing


Set wbOF1 = Nothing
Set wbB = Nothing
Set wbF = Nothing
'Application.ScreenUpdating = True
Updateprogress , 1, "Done...", True

MsgBox "Done!"

End Sub
Sub Updateprogress(Optional ByVal strDesc As Variant, Optional ByVal dblPercent As Variant, Optional ByVal strTitle As Variant, Optional ByVal bShow As Boolean)

Dim dblBarWidth As Double
dblBarWidth = 354

If blShowProgress = True Then
If Not IsMissing(strTitle) Then
    frmprogress.lblTitle = strTitle
End If

If Not IsMissing(strDesc) Then
    frmprogress.lblDesc = strDesc
End If

If IsNumeric(dblPercent) Then
    frmprogress.lbl.Width = dblPercent * dblBarWidth
End If
        
frmprogress.lblnow.Caption = "     Time Now : " & Now

If bShow = True Then
    frmprogress.Show
Else
    DoEvents
'    frmprogress.Repaint
End If

End If
End Sub
Function GetPrimaryGivingPlatform(id As Long, wb As Workbook) As String
Dim shtcalc As Worksheet
Dim Lastrow As Long
Dim x As Long
Dim shtb As Worksheet
Dim lngmost As Long
Dim strP As String
Dim c As Object
Set shtcalc = wb.Sheets("Calc")
shtcalc.Cells.ClearContents
shtcalc.Range("A1").Formula2 = "=unique(filter(B!C:C,B!A:A=" & id & "))"
Set c = shtcalc.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
If Not c Is Nothing Then
    Lastrow = c.Row
Else
    Lastrow = 0
End If
If IsError(shtcalc.Range("A1")) Then
    GetPrimaryGivingPlatform = "None"
    Exit Function
End If

If Lastrow = 0 Then
    GetPrimaryGivingPlatform = "None"
    Exit Function
End If

For x = 1 To Lastrow
    shtcalc.Range("B" & x).Formula2 = "=COUNTIF(B!C:C,""="" & A" & x & ")"
Next x
lngmost = shtcalc.Range("B1")
strP = shtcalc.Range("A1")
For x = 2 To Lastrow
    If shtcalc.Range("B" & x) > lngmost Then
        lngmost = shtcalc.Range("B" & x)
        strP = shtcalc.Range("A" & x)
    End If
Next x
GetPrimaryGivingPlatform = strP

        



End Function
Function GetPrimaryGivingPlatform2(id As Long, wb As Workbook) As String
Dim shtcalc As Worksheet
Dim Lastrow As Long
Dim x As Long
Dim shtb As Worksheet
Dim lngmost As Long
Dim strP As String
Dim c As Object
Set shtcalc = wb.Sheets("Calc")
shtcalc.Cells.ClearContents
shtcalc.Range("A1").Formula2 = "=sort(unique(filter(B!I:C,B!A:A=" & id & ")),7,1)"
Set c = shtcalc.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
If Not c Is Nothing Then
    Lastrow = c.Row
Else
    Lastrow = 0
End If
If IsError(shtcalc.Range("A1")) Then
    GetPrimaryGivingPlatform2 = "None"
    Exit Function
End If

If Lastrow = 0 Then
    GetPrimaryGivingPlatform2 = "None"
    Exit Function
End If



    
    

GetPrimaryGivingPlatform2 = shtcalc.Range("G1")

        



End Function






Sub ClearAll()
ActiveSheet.Range("B6") = vbNullString
ActiveSheet.Range("B9") = vbNullString
ActiveSheet.Range("B12") = vbNullString

End Sub
Sub GetRFM()
Dim c As Object
Dim lastrowPivot As Long
Dim shtPivot As Worksheet
Dim shtO As Worksheet
Set shtO = Sheets("Output")
Set shtPivot = Sheets("Pivot")


    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Output!R1C1:R1048576C17", Version:=8).CreatePivotTable TableDestination:= _
        "Pivot!R1C1", TableName:="PivotTable1", DefaultVersion:=8
    With shtPivot.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With shtPivot.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    shtPivot.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With shtPivot.PivotTables("PivotTable1").PivotFields("Constituent ID")
        .Orientation = xlRowField
        .Position = 1
    End With
    shtPivot.PivotTables("PivotTable1").PivotSelect "'Constituent ID'[All]", _
        xlLabelOnly + xlFirstRow, True
    shtPivot.PivotTables("PivotTable1").AddDataField shtPivot.PivotTables( _
        "PivotTable1").PivotFields("Recency Criteria"), "Max of Recency Criteria", _
        xlMax
    shtPivot.PivotTables("PivotTable1").AddDataField shtPivot.PivotTables( _
        "PivotTable1").PivotFields("Frequency Criteria"), "Sum of Frequency Criteria", _
        xlSum
    shtPivot.PivotTables("PivotTable1").AddDataField shtPivot.PivotTables( _
        "PivotTable1").PivotFields("Monetary Criteria"), "Sum of Monetary Criteria", _
        xlSum
    With shtPivot.PivotTables("PivotTable1").PivotFields( _
        "Max of Recency Criteria")
        .NumberFormat = "m/d/yyyy"
    End With
    
    
    ''get lastrow on pivot
    Set c = shtPivot.Range("A:A").Find(What:="*", after:=[A1], LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If Not c Is Nothing Then
        lastrowPivot = c.Row
    End If
    
    shtPivot.Range("E1") = "Latest Donation Date"
    shtPivot.Range("F1") = "# Donations"
    shtPivot.Range("G1") = "$ Donations"
    shtPivot.Range("H1") = "Recency Score"
    shtPivot.Range("I1") = "Frequency Score"
    shtPivot.Range("J1") = "Monetary Score"
    shtPivot.Range("K1") = "RFM Score"
    shtPivot.Range("E2") = "=MAXIFS(B:B,A:A,""="" & $A2)"
    shtPivot.Range("E2").NumberFormat = "mm/dd/yyyy"
    shtPivot.Range("F2") = "=SUMIF(A:A,""="" & $A2,C:C)"
    shtPivot.Range("G2") = "=SUMIF(A:A,""="" & $A2,D:D)"
    shtPivot.Range("G2").NumberFormat = "$#,###,###,###.00"
    shtPivot.Range("H2") = "=IFERROR(RANK.EQ(E2,INDEX(A:G,0,COLUMN(E2)),1)/COUNT(INDEX(A:G,0,COLUMN(E2)))*10,0)"
    shtPivot.Range("I2") = "=IFERROR(RANK.EQ(F2,INDEX(A:G,0,COLUMN(F2)),1)/COUNT(INDEX(A:G,0,COLUMN(F2)))*10,0)"
    shtPivot.Range("J2") = "=IFERROR(RANK.EQ(G2,INDEX(A:G,0,COLUMN(G2)),1)/COUNT(INDEX(A:G,0,COLUMN(G2)))*10,0)"
    shtPivot.Range("K2") = "=Sum(H2:J2)"
    shtPivot.Range("H2:K2").NumberFormat = "#.00"
    
    shtPivot.Range("E2:K2").Copy shtPivot.Range("E3:K" & lastrowPivot)

    
End Sub








