Attribute VB_Name = "JustCode2_UpdateCodeSharedPath"
Option Explicit
''https://support.microfocus.com/kb/doc.php?id=7021399
'Managing Update VBA for Multiple Users
'------------------------------------------------
'Tools > References> select the Microsoft Visual Basic for Applications Extensibility
'The settings file containing this code can now be distributed to end users, and each time it is opened,
'an event will automatically update the local VBA project’s SharedMacroCode module with a new version retrieved
'from Z:\SharedMacro\SharedMacroCode.bas.
'If you provide a group of users a file that includes macro code and later it becomes necessary to update this macro code,
'replacing the file the personal data and settings would be lost.
'To avoid this problem and to simplify administration of enable everyone to receive updates automatically on a regular basis;
'for example, each time a session is opened.

'' Method3: Push new code from a shared path on local network to orignial users file (VBA)
Public Sub UpdateCodeLocalpath()
Const myPath As String = "X:\SharedMacroCode\JustCode_SomeCodeToReplace.bas"
Const ModuleName As String = "JustCode_SomeCodeToReplace"
'    On Error Resume Next
    
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
MsgBox "Could not update VBA code from: " & myPath & "Sub UpdateCodeLocalpath"
End Sub


