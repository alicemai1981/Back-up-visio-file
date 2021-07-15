VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'use system api to open Windows file dialog, source:https://stackoverflow.com/questions/33762039/open-a-filedialog-in-visio-vba
'https://www.exceltip.com/files-workbook-and-worksheets-in-vba/select-folder-names-using-vba-in-microsoft-excel.html
'https://jkp-ads.com/Articles/apideclarations.asp

Private Declare PtrSafe Function GetFileOpen Lib "comdlg32.dll" _
Alias "GetOpenFileNameA" (OFD As OpenFileName) As Boolean

Private Type OpenFileName

  lStructSize As Long
  hwndOwner As LongPtr
  hInstance As LongPtr
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As LongPtr
  lpTemplateName As String
End Type

Private Type BROWSEINFO
    hOwner As LongPtr
    pidlRoot As LongPtr
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As LongPtr
    lParam As LongPtr
    iImage As Long
End Type

Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal Psz As String) As Boolean
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr


Sub OpenFile(ByRef FilePath As String, ByRef cancelled As Boolean)
Dim OpenFile As OpenFileName
Dim sFilter As String
Dim lReturn As Long

OpenFile.lStructSize = LenB(OpenFile)
sFilter = "All Files(*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar

OpenFile.lpstrFilter = sFilter
OpenFile.nFilterIndex = 1
OpenFile.lpstrFile = String(257, 0)
OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
OpenFile.lpstrFileTitle = OpenFile.lpstrFile
OpenFile.nMaxFileTitle = OpenFile.nMaxFile
OpenFile.lpstrInitialDir = ThisDocument.Path

OpenFile.lpstrTitle = "Find the data source"
OpenFile.flags = OFN_EXPLORER
lReturn = GetFileOpen(OpenFile)

If lReturn = 0 Then
  cancelled = True
  FilePath = vbNullString
Else
  cancelled = False
  FilePath = Trim(OpenFile.lpstrFile)
  FilePath = Replace(FilePath, Chr(0), vbNullString)
End If

End Sub

Function GetFolderName() As String
Dim BInfo As BROWSEINFO
Dim ItemLoc As LongPtr
Dim Path As String
Dim OK As Boolean
Dim pos As Integer

BInfo.pidlRoot = 0
BInfo.ulFlags = &H10 Or &H40 Or &H4000 Or &H80 Or &H8000   'flag info https://docs.microsoft.com/en-us/windows/win32/api/shlobj_core/ns-shlobj_core-browseinfoa
ItemLoc = SHBrowseForFolder(BInfo) 'display dialog,specifies the location of the selected folder relative to the root of the namespace.
Path = Space(512)
OK = SHGetPathFromIDList(ItemLoc, Path)
If OK Then
    pos = InStr(Path, Chr(0))
    GetFolderName = Left(Path, pos - 1)
Else
    GetFolderName = ""
End If

End Function





