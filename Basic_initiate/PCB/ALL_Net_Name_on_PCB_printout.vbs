
option Explicit
Dim pcbApp
Set pcbApp = Application
' Set pcbApp = GetObject (,"MGCPCB.Application")
Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
   Call Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
'    Call Scripting.AddTypeLibrary("MGCPCB.Application")

   Call Scripting.AddTypeLibrary("Scripting.FileSystemObject")

   
    ' 여기에 코드를 작성 시작!
 
    ' 퍼포먼스 향상을 위해 잠시 PCB를 멈춤
    pcbApp.LockServer     


    Dim NetColl, NetObj, NetName, NetCount, i:  i=0

    ' Nets obejct 가져와서 Net Collection 생성
    Set netColl = pcbDoc.Nets

    ' PCB에 있는 전체 Net 갯수
    NetCount = NetColl.Count

    ' 각하나의 Net이름 출력
    For Each netObj In NetColl
        NetName = netObj.Name
        i = i + 1
        Call AppendOutput("User Input", "Net Name : " & NetName)
        Call AppendOutput("User Input", "Number " & i & " of total " & netCount)
        Call AppendOutput("User Input", "===========")

    Next

    ' 코드가 다 실행 된 후 다시 PCB를 재가동 (코드 성능 향상)
    pcbApp.UnlockServer
    


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
