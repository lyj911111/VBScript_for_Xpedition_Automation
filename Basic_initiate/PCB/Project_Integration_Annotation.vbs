'
' Project Integration 접근 Sync Annotate을 시킴
'


option Explicit
Dim pcbApp
Set pcbApp = Application
' Set pcbApp = GetObject (,"MGCPCB.Application")
Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Call Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    'Call Scripting.AddTypeLibrary("MGCPCB.Application")
    Call Scripting.AddTypeLibrary("Scripting.FileSystemObject")


    ' 여기에 코드를 작성 시작!
    Dim PrjIntObj
    Set PrjIntObj = pcbDoc.ProjectIntegration

    ' PCB 디자인의 Forward Annotate 를 시킴
    PrjIntObj.ForwardAnnotate

    Set PrjIntObj = Nothing

Else
    Msgbox("Could not validate the server. Exiting program.")
End If


'---------------------------------------
' 여기에 sub루틴 함수 작성
'---------------------------------------



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
    ' Ask Expedition document for the key
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
