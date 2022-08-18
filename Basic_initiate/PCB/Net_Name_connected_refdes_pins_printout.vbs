'
' PCB상 Net Name 을 기준으로 어떤 부품(refdes) 의 몇번째 핀(pin)에 붙어있는지 확인
'
'


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

    Dim fso,FileName,LogFile,CurProjPath,lfileName,Win
    Dim PinColl, PinCount, RefDes, CompFromPin, NetPin
    Dim netColl, NetName, netObj
   ' Create the FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")

   ' 유저 입력을 받아서 파일명 생성함. default 자동완성 이름은 "NetRefdesPin.txt" 
   FileName = InputBox ("Output File Name: ","Output File","NetRefdesPin.txt")
   ' Acquire the project path of the open document
   CurProjPath = pcbDoc.Path   ' CurProjPath is set to project directory
   ' Create variable whose value is the complete path/filename for report
   lfileName = CurProjPath & FileName  '  "FileName" input by user
   ' Create the report file in the project directory 파일 생성
   Set LogFile = fso.CreateTextFile(lfileName, True)


   ' Document the beginning of the report with a note and column headings text파일 시작부 작성 (해더)
   LogFile.Write("This is a list of all the net/pin connections." & vbCrLf)
   LogFile.WriteBlankLines(4)
   LogFile.Write("  Netname       Refdes-pin" & vbCrLf)
   

   pcbApp.LockServer        ' Always use with large collections - runs faster

   ' Net 이름 기준으로 붙어있는 Pin들을 색출!
   Set netColl = pcbDoc.Nets   ' Create the collection of all the nets in the design
   For Each netObj In netColl  ' Process each net in the design
        NetName = netObj.Name   ' Acquire the Net name
        Set PinColl = netObj.pins  ' Create the collection of pins
        pincount=0                          ' Initialize the variable for pins/net summary
        For Each netPin In PinColl     ' Processes each pin in the design
            Set CompFromPin = NetPin.Component   ' Acquires comp. pin attached to net
            RefDes = CompFromPin.Name         ' Acquires the RefDes
            LogFile.Write ("  " & NetName & "          " & Refdes & "-" & NetPin & vbCrLf)
            pincount = pincount + 1          ' Increment pincounter for pins/net summary
        Next ' For Each netPin
        LogFile.Write("Total number of pins for net " & NetName & " is " & pincount & vbCrLf)
        LogFile.Write("=============================")
        LogFile.WriteBlankLines(2)
        LogFile.Write("  Netname       Refdes-pin" & vbCrLf)
   Next
   LogFile.WriteBlankLines (2)
   LogFile.Write("Report Complete.")
   LogFile.Close

   ' 성공했음을 알림
   MsgBox "The Logfile closed successfully.",,"File Close Confirmed"

   pcbApp.UnlockServer

   ' 끝나고 자동으로 메모장 실행
   Set Win =CreateObject("WScript.shell")   ' Allows launching of other executables
   Win.run "notepad.exe  " &  lfileName   ' Opens lfileName in Notepad
    


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
