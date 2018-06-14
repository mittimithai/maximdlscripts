' Assumes object info is in clipboard from TheSky and writes focus star coords to perisistent dictionary

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


starinfo= CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")
If IsNull(QuickClip) Then starinfo= ""

lines = Split(starinfo,vbCrLf)
coords = ""
ra_str = "ERROR"
dec_str = "ERROR"

Dim Util
Set Util = CreateObject( "ASCOM.Utilities.Util" )

For Each line in lines
 If Instr(1, line,"Equatorial:") Then
    coords = line
    tokens = Split(coords)
    ra_str = tokens(2) & tokens(3) & tokens(4) & tokens(5)
    dec_str = tokens(8)
 End If
Next

result = MsgBox ("Selected focus star at coordinates: " & ra_str & " " & dec_str & vbCrLf & "Ready to slew and focus?", vbYesNo, "Focus Star")

If result = vbNo Then
 Error = "Cancelled"
 wscript.Quit
End If


Include("C:\ccdaq\script\transdictionary.vbs")
FocusStarFile="C:\ccdaq\script\focusstar.txt"

Dim transDict
Set transDict = New TransDictionary
transDict.LoadLog(FocusStarFile)
transDict.SetValue "focusstar_ra",ra_str
transDict.SetValue "focusstar_dec",dec_str

transDict.Commit

Dim objShell
Set objShell = wscript.CreateObject("WScript.Shell")

' objShell.Run "C:\ccdaq\script\fmx_focus.vbs" 

' Using Set is mandatory
Set objShell = Nothing