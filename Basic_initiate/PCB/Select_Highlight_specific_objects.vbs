'
' Filter 사용예 PCB상의 특정 데이터 필터
' 특정 Refdes로 필터 하여 PartNumber와 Refdes 만 뽑아서
' 모두 출력
' 특정 오브젝트를 마치 마우스 드레그 한것 처럼 선택 함!~
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

    Dim NetName, NetObj, CollViaTotal, ViaCountTotal, CollVia, ViaCount, viaPstkObj

    ' Via를 모두 변경할 "NET" 이름을 User로 부터 입력받음
    NetName = InputBox ("Modify vias for Net: ","User Prompt")

    ' 유저가 준 Net 이름으로 Net object 생성 (해당 net name이 PCB안에 있어야 함!)
    Set NetObj = pcbDoc.FindNet(NetName)
    ' 그래픽상에서 해당 User가 준 Net object를 선택함
    NetObj.Selected = True

    ' PCB전체의 via collection을 가져옴
    Set CollViaTotal = pcbDoc.vias   ' Create a collection of all the vias
    ViaCountTotal = CollViaTotal.count  ' Count of all the vias in the design
    MsgBox "Total Via : " & ViaCountTotal

    ' 현재 하이라이트 시킨 via의 collection을 가져옴
    Set CollVia = NetObj.vias       ' Create a collection of all the vias on the Net
    ViaCount = CollVia.count       ' Count of all the vias on the net
    MsgBox "Selected Via : " & ViaCount

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
