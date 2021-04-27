Sub OutputCSV_version2()


'Check for output folder and create if required
Dim strFolderName As String
Dim strFolderExists As String
    strFolderName = Application.ActiveWorkbook.Path & "\CSV_Uploads\"
    strFolderExists = Dir(strFolderName, vbDirectory)
    If strFolderExists = "" Then
        MsgBox ("Creating output " & Application.ActiveWorkbook.Path & " \ CSV_Uploads \ ")
        MkDir Application.ActiveWorkbook.Path & "\CSV_Uploads\"
    Else
    End If

'start copy

Dim Rng As Range
Dim WorkRng As Range
Dim xFile As Variant
Dim xFileString As String
On Error Resume Next

Dim rowstocopy As Integer
Dim Rng2 As Range
Dim myfile As String
myfile = Application.ActiveWorkbook.Path & "\CSV_Uploads\Output--" & Format(Now, "yyyy-mm-dd--hh-mm-ss") & ".csv" 'Output file name and path

rowstocopy = Evaluate("=max(IF(len(C$1:C$250)>0,ROW(C$1:C$250),0))") ' row count dynamic on col C
Set Rng2 = Range("A1:BX" & rowstocopy) 'Hard code of starting cell and col to copy to CSV file

Set WorkRng = Application.Selection
Set WorkRng = Rng2


Application.ActiveSheet.Copy
Application.ActiveSheet.Cells.Clear
WorkRng.Copy Application.ActiveSheet.Range("A1")
Set xFile = CreateObject("Scripting.FileSystemObject")
ActiveWorkbook.SaveAs Filename:=myfile, FileFormat:=xlCSV, CreateBackup:=False
ActiveWorkbook.Close savechanges:=False


MsgBox (rowstocopy & " Rows Copied to " & Chr(13) & Chr(10) & myfile) 'Confirmation Box



End Sub