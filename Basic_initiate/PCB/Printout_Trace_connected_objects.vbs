'
'
' 선택된 (highlighted) Trace가 연결된 부분을 display 해줌.
' Trace가 Obstruct가 있다면 굳이 Mouting Hole이나 Fiducial Marker와 연결될 일 없으므로
' 아래 출력하는 부분은 생략함. 그래서 refdes나, via, pin number만 출력해놓게 했음
'

Option Explicit

Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")
	
	Dim colTraces, oTrace, oConn, oConns, Type2
	
	Set colTraces = pcbDoc.Traces(epcbSelectSelected,0)

	On Error Resume Next
	For Each oTrace In colTraces

		For Each oConn In oTrace.ConnectedObjects 
            
			AppendOutput "test", "------------------------------------------------"
			AppendOutput "test", "oTrace.Net = " & oTrace.Net 
			
			Dim ptype
			Dim pn : pn = "--"
			Dim refdes : refdes = "--"
			Select Case oConn.Type
				Case epcbPadstackObjectPin
					ptype = " Selected Type is -> pin"
					pn = oConn.Name
					refdes = oConn.Component.Refdes
				Case epcbPadstackObjectVia
					ptype = " Selected Type is -> via"
				Case epcbPadstackObjectFiducial
					ptype = " Selected Type is -> fiducial"
				Case epcbPadstackObjectMountingHole
					ptype = " Selected Type is -> mounting hole"
				Case epcbPadstackObjectToolingHole
					ptype = " Selected Type is -> tooling hole"
				Case epcbPadstackObjectShearingHole
					ptype = " Selected Type is -> shearing hole"
				Case epcbPadstackObjectMultVia
					ptype = " Selected Type is -> multi via"
			End Select			
			
			AppendOutput "test", "oConn.Type : " & ptype
            AppendOutput "test", "pin Number= " & pn 
            AppendOutput "test", "refdes= " & refdes
		Next
	Next	
	On Error GoTo 0
    AppendOutput "test", "***********************************************"

	
Else
    Msgbox("Could not validate the server. Exiting program.")
End If





'=========================================================================
' Message Window Output
'=========================================================================
Function AppendOutput(sOutputTab, str)
	Dim mnu, OutputControl, objTab
	Set mnu = Gui.CommandBars("Document Menu Bar").Controls("&View").Controls("Message Window")

	If mnu.Checked = False Then
		Call Gui.ProcessCommand(33125)
	End If

	Set OutputControl = Addins.Item("Message Window").Control
	Set objTab = OutputControl.AddTab(sOutputTab)
	Call objTab.Activate
	Addins("Message Window").Control.AddTab(sOutputTab).AppendText (str & vbCrLf)
End Function

Function ClearOutputWindow(sOutputTab)
	Addins("Message Window").Control.AddTab(sOutputTab).Clear
End Function

'---------------------------------------
' Begin Validate Server Function
'---------------------------------------
Private Function ValidateServer(doc)
    
    Dim key, licenseServer, licenseToken

    ' Ask Expedition뭩 document for the key
    key = doc.Validate(0)

    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)

    ' Release license server
    Set licenseServer = nothing

    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    doc.Validate(licenseToken)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If

End Function
'---------------------------------------
' End Validate Server Function

