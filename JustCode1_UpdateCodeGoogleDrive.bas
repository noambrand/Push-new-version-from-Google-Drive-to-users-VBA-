Attribute VB_Name = "JustCode1_UpdateCodeGoogleDrive"
Option Explicit
''Method2: Push new code from Google Drive to original users file (VBA)
Public myPath As String
Const ModuleName As String = "JustCode_SomeCodeToReplace"

Sub RunDownloadCODEGoogleDriveVersion()
Dim response As String
''myOriginalURL - The original google drive URL path (Before modifications of UrlLeft & FileID & UrlRight)
' filetypeNewVersion - doc/ drive (see CASE in filetypeNewVersion)
' OpenFolderPath- open new file? the first time false, the second time can be true.
Call DownloadGoogleDrive(PushVersion.Range("A5"), "doc", False)
Call TextIORead(PushVersion.Range("C5"))  ' If a newer version is avialable it will return MostUpdated=FALSE as global variable
''If MostUpdated=FALSE Run DownloadGoogleDrive to updated workbook, otherwise do nothing.
If Not MostUpdated Then
    PushVersion.Range("A6") = newURL
' if Downloads aleardy has the file delete it so the downloaded file won't be renamed to filename(1)
    myPath = Environ$("USERPROFILE") & "\Downloads\" & ModuleName & ".bas"
    Kill myPath
    ' open browser with google drive download path
    ThisWorkbook.FollowHyperlink Address:=newURL
' User has to Download the BAS file manually to his Downloads folder
    response = MsgBox("First confirm download BAS file to your download folder " & vbCrLf & _
    "then Press 'OK'", vbOKCancel + vbQuestion)
    If response = vbOK Then UpdateCodeGoogleDrive
End If
End Sub

'' Update code from a location on Google drive
Public Sub UpdateCodeGoogleDrive()
    On Error GoTo skip
    'include reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
    Dim vbproj As VBProject
    Dim vbc As VBComponent
    Set vbproj = ThisWorkbook.VBProject
    
    'Error will occur if component with this name is not in the project, so this will help avoid the error
    Set vbc = vbproj.VBComponents.Item(ModuleName)
    If Err.Number <> 0 Then
        Err.Clear
        vbproj.VBComponents.Import myPath
        If Err.Number <> 0 Then GoTo skip
    Else
        'no error - vbc should be valid object
        'remove existing version first before adding new version
        vbproj.VBComponents.Remove vbc
        vbproj.VBComponents.Import myPath
        If Err.Number <> 0 Then GoTo skip
    End If
    
Exit Sub
skip:
MsgBox "Could not update VBA code from: " & myPath & "Sub UpdateCodeGoogleDrive"
End Sub

