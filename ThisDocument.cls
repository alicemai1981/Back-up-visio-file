VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub AutoBackUp()
Attribute AutoBackUp.VB_ProcData.VB_Invoke_Func = "m"
Dim fd As New FileDialog
Dim UserFolderPath As String
Dim BackupName As String
Dim fso As Object
Dim FilesInBackup As Object
Dim EarlestVersion As Object
Dim BackupObj As Visio.Document

ThisDocument.Save
UserFolderPath = fd.GetFolderName
RandomString = CreateRandomNumb()
BackupName = Replace(RandomString & ThisDocument.Name, "vsdm", "vsd", 1, 1)
BackupNumb = CountFiles(UserFolderPath)
FullBackupName = UserFolderPath & "\" & BackupName

Set BackupObj = Documents.Add(ThisDocument.Path & ThisDocument.Name)
If BackupNumb >= 10 Then
    Set EarlestVersion = FindOldest(UserFolderPath)
    EarlestVersion.Delete
End If
    BackupObj.SaveAs (FullBackupName)
    BackupObj.Close

End Sub

Function CreateRandomNumb() As String
 
SysTime = Time
Dim i As Long
StrHash = &H65D5BAAA

For i = 1 To Len(SysTime)
    StrHash = ((StrHash + AscW(Mid$(SysTime, i, 1))) Mod 69208103) * 31& 'source: https://stackoverflow.com/questions/14717526/vba-hash-string
Next

CreateRandomNumb = CStr(StrHash)

End Function

Function CountFiles(FilesPath As String) As Integer 'source: https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_other-mso_2007/how-to-find-number-of-files-in-a-folder-via-vba/87c650de-c0b3-46b0-b7d2-284e0b3b328f
Dim fso As Object
Dim objFiles As Object
Dim FilePath As String
Dim count As Integer

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFiles = fso.GetFolder(FilesPath).Files
count = objFiles.count
CountFiles = count

End Function

Function FindOldest(UserFolderPath) As Object
Dim FilesInBackup As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set FilesInBackup = fso.GetFolder(UserFolderPath).Files

Static EarlestCreatedDate, EarlestFile As Object
EarlestCreatedDate = CDate(Time)

For Each f In FilesInBackup
    If f.DateCreated < EarlestCreatedDate Then
        EarlestCreatedDate = f.DateCreated
        Set EarlestFile = f
   'Else
    'Debug.Print f.DateCreated, EarlestCreatedDate
    End If
Next f

Set FindOldest = EarlestFile
End Function

