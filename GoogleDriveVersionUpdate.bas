Attribute VB_Name = "GoogleDriveVersionUpdate"
Option Explicit
' Local files Version update by VBA (VBA is contained in the original file you distribute).
' Verify if an updated version of the file is available and download it.
' RunDownloadGoogleDriveVersion is called evreytime the workbook is opened and quietly
' downloads a text file from a public GoogleDrive folder depending on the content of the
' text file the a new workbook path will be used to downloaded the new version.
' The Google doc file on google drive will be delemiatated by  ";" in the format:
' [Newversion number] ; [Google drive link] ; [WhatsNewInVersion a meassage to dispaly to the user] e.g.:
' 8;https://drive.google.com/file/d/[FileID]/view?usp=sharing; A new version is available.

' Method1: Push new file version from Google Drive to users (VBA)
Public filetypeNewVersion As String
Public newURL As String
Public MostUpdated As Boolean
Public WhatsNewInVersion As String
Public versionNumINT As Long

Sub RunDownloadGoogleDriveVersion()
Call DownloadGoogleDrive(PushVersion.Range("A3"), "doc", False)
Call TextIORead(PushVersion.Range("C3")) ' If a newer version is avialable it will read its path on Google drive
If Not MostUpdated Then
    PushVersion.Range("A4") = newURL
    Call DownloadGoogleDrive(newURL, PushVersion.Range("B4"), True)
End If
End Sub
  
' myOriginalURL - The original google drive URL path (Before modifications of UrlLeft & FileID & UrlRight)
' filetypeNewVersion - doc/ drive (see CASE in filetypeNewVersion)
' OpenFolderPath- open new file? the first time use False, the second time you can choose True.

Sub DownloadGoogleDrive(myOriginalURL As String, filetypeNewVersion As String, OpenFolderPath As Boolean)
Dim myURL As String
Dim FileID As String
Dim xmlhttp As Object
Dim name0 As Variant
Dim FolderPath As String
Dim FilePath As String
Dim oStream As Object
Dim wasDownloaded As Boolean
Application.ScreenUpdating = False

Dim UrlLeft As String
Dim UrlRight As String
Select Case filetypeNewVersion
    Case "doc" 'for Google doc or Google Sheets
        UrlLeft = "https://docs.google.com/document/d/"
        UrlRight = "/export?format=txt"
    Case "drive" 'for EXCEL, PDF, WORD, ZIP etc saved in Google Drive
        UrlLeft = "http://drive.google.com/u/0/uc?id="
        UrlRight = "&export=download"
    Case Else
        MsgBox "Wrong file type", vbCritical
        End
End Select

''URL from share link or Google sheet URL or Google doc URL
'' myOriginalURL = "https://drive.google.com/file/d/..." ''myVersionUpdateWarning
''https://drive.google.com/drive/folders/....
'' Credit to Florian Lindstaedt: https://www.linkedin.com/pulse/20140608044541-54506939-how-to-recall-an-old-excel-spreadsheet-version-control-with-vba
    FileID = Split(myOriginalURL, "/d/")(1) ''split after "/d/"
    FileID = Split(FileID, "/")(0)  ''split before single "/"
    myURL = UrlLeft & FileID & UrlRight

        Set xmlhttp = CreateObject("winhttp.winhttprequest.5.1")
        xmlhttp.Open "GET", myURL, False  ', "username", "password"
        xmlhttp.Send
        
On Error Resume Next
        name0 = xmlhttp.getResponseHeader("Content-Disposition")
    If Err.Number = 0 Then
            If name0 = "" Then
                  MsgBox "file name not found", vbCritical
                  Exit Sub
             End If
                  Debug.Print name0
                  name0 = Split(name0, "=""")(1) ''split after "=""
                  name0 = Split(name0, """;")(0)  ''split before "";"
                  Debug.Print name0
                  Debug.Print FilePath
    End If
        
   If Err.Number <> 0 Then
         Err.Clear
         Debug.Print xmlhttp.responseText
        ''<a href="/open?id=FileID">JustCode_CodeUpdate.bas</a>
         name0 = xmlhttp.responseText
         name0 = ExtractPartOfstring(name0)
    End If
On Error GoTo 0

    FolderPath = ThisWorkbook.path
    FilePath = FolderPath & "\" & name0
 ''This part is does the same as Windows API URLDownloadToFile function(no declarations needed)
 On Error GoTo skip
    If xmlhttp.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
                .Open
                .Charset = "utf-8"
                .Type = 1  'Binary Type
                .Write xmlhttp.responseBody
                .SaveToFile FilePath, 2 ' 1 = no overwrite, 2 = overwrite
                .Close
        End With
    End If
    
 Application.ScreenUpdating = True
 
  If FileExists(FilePath) Then
              wasDownloaded = True
              ''open folder path location to look at the downloded file
             If OpenFolderPath Then Call Shell("explorer.exe" & " " & FolderPath, vbNormalFocus)
        Else
              wasDownloaded = False
              MsgBox "failed", vbCritical
  End If
  Exit Sub
skip:
   MsgBox "Tried to download file with same name as current file," & vbCrLf & _
          "check in google docs the version number and link are correct", vbCritical
End Sub


'TextIORead opens a text file, retrieving some text, closes the text file.
Sub TextIORead(TXTname As String)
On Error GoTo skip
  Dim sFile As String
  Dim iFileNum As Long
  Dim sText As String
  Dim versionNum As String
  sFile = ThisWorkbook.path & "\" & TXTname
  
  If Not FileExists(sFile) Then
        MsgBox "version download doc file not found", vbCritical
        End
  End If

'For Input - extract information. modify text not available in this mode.
'FreeFile - supply a file number that is not already in use. This is similar to referencing Workbook(1) vs. Workbook(2).
'By using FreeFile, the function will automatically return the next available reference number for your text file.
  iFileNum = FreeFile
  Open sFile For Input As iFileNum
  Input #iFileNum, sText
  Close #iFileNum
  
versionNum = Split(sText, ";")(0)
versionNum = Replace(versionNum, "ï»¿", "") ''junk caused by the UTF-8 BOM that can't be changed when downloading from google docs
'versionNum = RemoveJunk(versionNum)
versionNumINT = VBA.CLng(versionNum)
newURL = Split(sText, ";")(1)
WhatsNewInVersion = Split(sText, ";")(2) ' split by semi-colons but also "," splits it!!!!?!

MostUpdated = CheckVersionMostUpdated(versionNum, newURL)
''Comment out for tests- sFile is just a temporary file that the user doesn't need and can just be deleted.
'Kill sFile
Exit Sub
skip:
MsgBox "The updated file was not found, please contact the developer for the new version", vbCritical
End Sub

Function CheckVersionMostUpdated(ByVal versionNum As String, ByVal newURL As String) As Boolean
Dim wkbVersion As String
Dim wkbVersionINT As Long
Dim response As String
wkbVersion = ThisWorkbook.Name
wkbVersion = Split(wkbVersion, "_")(1)
wkbVersion = Split(wkbVersion, ".")(0)
wkbVersionINT = VBA.CLng(wkbVersion)
Debug.Print wkbVersion
CheckVersionMostUpdated = True
If versionNumINT > wkbVersionINT Then
''Hebrew Display problems caused by the UTF-8 BOM:  https://www.w3.org/International/questions/qa-utf8-bom.en.html
MsgBox WhatsNewInVersion, vbInformation
' Download new version?
    response = MsgBox("This workook version: " & wkbVersion & vbCrLf & _
    "Available version: " & versionNum & vbCrLf & _
    "There is a newer version available, Download to the current file folder?", vbOKCancel + vbQuestion)
    If response = vbOK Then CheckVersionMostUpdated = False
    If response = vbCancel Then CheckVersionMostUpdated = True
    Else
    MsgBox "You have the most updated version", vbInformation
End If
End Function

Function FileExists(FilePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    FileExists = True
    If TestStr = "" Then
        FileExists = False
    End If
End Function

'' mystring= <a href="/open?id=1HYx4987q2dB1M1OEginG5dTnD2SIwsy-">JustCode_CodeUpdate.bas</a>
Function ExtractPartOfstring(ByVal mystring As String) As String
  Dim first As Long, second As Long
  second = InStr(mystring, "</a>")
  first = InStrRev(mystring, ">", second)
  ExtractPartOfstring = Mid$(mystring, first + 1, second - first - 1)
  Debug.Print ExtractPartOfstring
End Function

'Function RemoveJunk(ByVal sInp As String) As String
'    Dim idx As Long, ArrayAscii(255) As Variant
'    Dim i As Long
'  ''https://www.ascii-codes.com/cp862.html
'    'Define Array for ASCII codes for Non-Printable Characters
'For i = 0 To 31
'  ArrayAscii(i) = i
'Next
'
'For i = 128 To 255
'  ArrayAscii(i - 128 + 32) = i
'Next
'
'    'Loop Thru Each Element in Array & Verify whether Any Special Character appears in String
'    For idx = LBound(ArrayAscii) To UBound(ArrayAscii)
'        If InStr(sInp, Chr(ArrayAscii(idx))) Then
'            sInp = Replace(sInp, Chr(ArrayAscii(idx)), "")
'        End If
'    Next
'
'    'Return String After removing Junk Values
'    RemoveJunk = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(sInp))
'End Function
'
'
'
'''Display problems caused by the UTF-8 BOM:  https://www.w3.org/International/questions/qa-utf8-bom.en.html
'Sub WriteUTF8WithoutBOM()
'Const adReadAll = -1
'Const adSaveCreateOverWrite = 2 'Overwrites the file
'Const adTypeBinary = 1
'Const adTypeText = 2
'Const adWriteChar = 0
'Const adModeReadWrite = 3
'Const adLF = 10
'Const adWriteLine = 1
'    Dim UTFStream As Object
'    Set UTFStream = CreateObject("adodb.stream")
'    UTFStream.Type = adTypeText
'    UTFStream.Mode = adModeReadWrite
'    UTFStream.Charset = "UTF-8"
'    UTFStream.LineSeparator = adLF
'    UTFStream.Open
'    UTFStream.WriteText "This is an unicode/UTF-8 test.", adWriteLine
'    UTFStream.WriteText "First set of special characters: öãâëòâããâéâéâäåñüûú€", adWriteLine
'    UTFStream.WriteText "Second set of special characters: qwertzuiopõúasdfghjkléáûyxcvbnm\|Ä€Í÷×äðÐ[]í³£;?¤>#&@{}<;>*~¡^¢°²`ÿ´½¨¸0", adWriteLine
'
'    UTFStream.Position = 3 'skip BOM
'
'    Dim BinaryStream As Object
'    Set BinaryStream = CreateObject("adodb.stream")
'    BinaryStream.Type = adTypeBinary
'    BinaryStream.Mode = adModeReadWrite
'    BinaryStream.Open
'
'    'Strips BOM (first 3 bytes)
'    UTFStream.CopyTo BinaryStream
'
'    UTFStream.SaveToFile "C:\Users\noamb\Downloads\EXCEL_VBA_Noam2\adodb-stream2.txt", adSaveCreateOverWrite
'    UTFStream.flush
'    UTFStream.Close
'
'
'  Dim sFile As String
'  Dim iFileNum As Long
'  Dim sText As String
'
''  sFile = ThisWorkbook.path & "\myVersionUpdateWarning.txt"
'  sFile = "C:\Users\noamb\Downloads\EXCEL_VBA_Noam2\adodb-stream2.txt"
'  If Not FileExists(sFile) Then
'        MsgBox "version download doc file not found", vbCritical
'  End If
'
'  iFileNum = FreeFile
'
'      Open sFile For Input As iFileNum
'  Input #iFileNum, sText
'  Close #iFileNum
'Debug.Print sText
'Dim WhatsNewInVersion As String
'WhatsNewInVersion = Split(sText, ";")(3)
''Debug.Print WhatsNewInVersion
''    MsgBox WhatsNewInVersion
'End Sub

''https://www.labnol.org/internet/direct-links-for-google-drive/28356/
'' direct download links for Google Docs or Google Sheets


'https://www.linkedin.com/pulse/20140608044541-54506939-how-to-recall-an-old-excel-spreadsheet-version-control-with-vba
' file version - used for version update checker
'How to recall an old Excel spreadsheet / version control with VBA

' the file in google drive contains a simple version number and message that will be processed to see if updates are available
'Create a simple text file with the following content:
'2;no; You are using an old version. Version V2 is available now. Please contact John Doe (john@doe.com) for the update as soon as possible. This file will close now.
'Save this file in your google drive folder as myVersionUpdateWarning.txt.
' Then log in to the Dropbox web site and navigate to your file and open it to get the public link (chain symbol in the upper right corner).


''https://www.linkedin.com/pulse/20140608044541-54506939-how-to-recall-an-old-excel-spreadsheet-version-control-with-vba/
'Public Sub updateMessage()
'
''On Error GoTo err1:
'
'Dim aFileContent() As String
'Dim strLastLine As String
'Dim strContent() As String
'Dim iVersion As Long
'Dim strForceExit As String
'Dim strMessage As String
'
'Dim objHttp As Object
'Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
'Call objHttp.Open("GET", cUpdateInfoFile, False)
'Call objHttp.Send("")
'
'' complete file content is now in objHttp.ResponseText
'aFileContent = Split(objHttp.responsetext, vbLf)
'
'' only use last line of file:
'strLastLine = aFileContent(UBound(aFileContent))
'Debug.Print strLastLine
'' split by semi-colons:
'strContent = Split(strLastLine, ";")
'Debug.Print strLastLine
'' fill variables
''iVersion = CInt(strContent(0))
''strForceExit = strContent(1)
''strMessage = strContent(2)
'
'' compare to file version stored in global constant
'If cFileVersion < iVersion Then
'        If VBA.LCase(strForceExit) = "yes" Then
'            MsgBox strMessage, vbCritical, "Closing file now" 'the workbook will be forced to close
'            ThisWorkbook.Close
'        Else
'            MsgBox strMessage, vbExclamation 'user can continue to use this file
'        End If
'End If
'
''err1:
'Stop
'
'End Sub



' 'short version with FileSystemObject but not good for doc URL
' Dim fs As Object
' Set fs = CreateObject("Scripting.FileSystemObject")
''FileCopy GetSpecialFolder(vbDirGoogleDrive) & "seriall.xlsx", "C:/MyDownloads/seriall.xlsx"
'FileCopy fs.GetSpecialFolder(myURL) & "/" & name0, ThisWorkbook.path & "/" & name0






