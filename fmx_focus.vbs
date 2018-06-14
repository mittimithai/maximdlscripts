' Slews to focus star in focusstar dictionary focuses with FocusMax and slews back to object

Option Explicit



Dim TelescopeDriver
TelescopeDriver = "AstroPhysicsV2.Telescope"

Dim Util
Set Util = CreateObject( "ASCOM.Utilities.Util" )

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
Dim transDict
Set transDict = New TransDictionary
transDict.LoadLog("C:\ccdaq\script\focusstar.txt")

dim focusstar_ra 
focusstar_ra = Util.HMSToHours(transDict.Item("focusstar_ra"))
dim focusstar_dec
focusstar_dec = Util.DMSToDegrees(transDict.Item("focusstar_dec"))

Dim Tel
Set Tel = wscript.CreateObject(TelescopeDriver)

dim object_ra
dim object_dec

object_ra = Tel.RightAscension
object_dec = Tel.Declination

'wscript.echo "Current ra/dec:" & " " & object_ra & " " & object_dec

If Not Tel.CanSlew Then
  Error = "Telescope cannot slew"
  wScript.Quit 1
End If
                        

'wscript.echo "Slewing to focus star: " & focusstar_ra & " " & focusstar_dec
Tel.SlewToCoordinates focusstar_ra, focusstar_dec


Dim FMx 
Dim FMxFoc

Dim Position

Dim SysName

Dim camera

Set camera = Nothing



If camera Is Nothing Then
   
	Set camera = CreateObject( "MaxIm.CCDcamera" )

	camera.DisableAutoShutdown = True

	camera.LinkEnabled = True

	If Not camera.LinkEnabled Then

		Error = "Failed to connect to camera"

		wscript.echo Error      
		wscript.Quit
   
	End If



	' filter names may not be available on the first exposure

	' unless MaxIm CCD is allowed to initialize fully
   
	wscript.Sleep 100

End If



'wscript.echo "Stopping guider and setting filterwheel to L" 


camera.GuiderStop()


camera.Filter = 3

Set FMx = CreateObject("FocusMax.FocusControl")

Set FMxFoc = CreateObject("FocusMax.Focuser")



'FMx.AcquireStarEnable = False

'FMx.AcquireStarSolveEnable = False

'FMx.AcquireStarFinalPointingUpdate = False

'FMx.AcquireStarReturnSlewEnable = False



FMx.Focus()


'wscript.echo "Done focusing, slewing back to object: " & object_ra & "," & object_dec
Tel.SlewToCoordinates object_ra, object_dec


'wscript.echo "Starting guider" 
camera.GuiderAutoSelectStar = True
camera.GuiderExpose(2.0)


while camera.GuiderRunning
  wscript.sleep 2000
wend
camera.GuiderTrack(2.0)



Set FMx = Nothing
Set FMxFoc = Nothing
Set camera = Nothing