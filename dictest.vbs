Option Explicit

Sub Include(sInstFile)
    Dim f, s, oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If oFSO.FileExists(sInstFile) Then
        Set f = oFSO.OpenTextFile(sInstFile)
        s = f.ReadAll
        f.Close
        ExecuteGlobal s
    End If
    On Error Goto 0
    Set f = Nothing
    Set oFSO = Nothing
End Sub

Include("C:\ccdaq\script\transdictionary.vbs")

Const LogFile="C:\ccdaq\script\Log.txt"
Const CleanLogFile="C:\ccdaq\script\CleanLog.txt"

Dim transDict,rc
Set transDict = New TransDictionary
' If the log file does not exists, TransDictionary will create it
' otherwise it will read it and load the data containted into the
' dictionary
transDict.LoadLog(LogFile)
If transDict.Exists("Ran Count") Then
    rc=transDict.Item("Ran Count")
    rc=rc+1
Else
    rc=1
End If
WScript.Echo "This script had been run " & rc & " times"
' VBScript will naturally convert the number into a String
' when it is stored. This sort of thing will fail if VBScript
' does not know a good way of converting the value to a String
' in which case you will have to do this yourself
transDict.SetValue "Ran Count",rc

' Now the dictionary has the new value but it is not committed
' so we can roll it back
WScript.Echo "The dictionary has the value " & transDict.Item("Ran Count")
transDict.Rollback
WScript.Echo "After rollabck it has the value " & transDict.Item("Ran Count")
transDict.SetValue "Ran Count",rc
' This commits the change and writes it to disk
transDict.Commit
' This writes a clean log file
transDict.CreateCleanLog CleanLogFile