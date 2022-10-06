'==============================================================================================================================
' This material contains trade secrets or otherwise confidential information owned by Siemens Industry Software Inc. 
' or its affiliates (collectively, “SISW”), or its licensors. Access to and use of this information is strictly limited as set 
' forth in the Customer’s applicable agreements with SISW. This material may not be copied, distributed, or otherwise 
' disclosed outside of the Customer’s facilities without the express written permission of SISW, and may not be used 
' in any way not expressly authorized by SISW.

' This document is for information and instruction purposes. Siemens reserves the right to make changes in 
' specifications and other information contained in this publication without prior notice, and the reader should, in all cases, 
' consult Siemens to determine whether any changes have been made.

' No representation or other affirmation of fact contained in this publication shall be deemed to be a warranty or give rise 
' to any liability of Siemens whatsoever.

' SIEMENS MAKES NO WARRANTY OF ANY KIND WITH REGARD TO THIS MATERIAL INCLUDING, BUT NOT LIMITED 
' TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NON-
' INFRINGEMENT OF INTELLECTUAL PROPERTY. SIEMENS SHALL NOT BE LIABLE FOR ANY DIRECT, INDIRECT, 
' INCIDENTAL, CONSEQUENTIAL OR PUNITIVE DAMAGES, LOST DATA OR PROFITS, EVEN IF SUCH DAMAGES 
' WERE FORESEEABLE, ARISING OUT OF OR RELATED TO THIS PUBLICATION OR THE INFORMATION CONTAINED 
' IN IT, EVEN IF SIEMENS HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

' TRADEMARKS: The trademarks, logos and service marks ("Marks") used herein are the property of Siemens or other parties. 
' No one is permitted to use these Marks without the prior written consent of Siemens or the owner of the Marks, as applicable. 
' The use herein of third party Marks is not an attempt to indicate Siemens as a source of a product, but is intended to indicate 
' a product from, or associated with, a particular third party. A list of Siemens' trademarks may be viewed at: www.plm.automation.
' siemens.com/global/en/legal/trademarks.html and mentor.com/trademarks.
'==============================================================================================================================
' Last Update: 2022/7/29
' Created by WJ

Option Explicit

Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument


' 폴더 생성을 위함
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' 윈도우 Shell에 접근하기 위함
Dim Win
Set Win = CreateObject("WScript.shell")

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")

    '===여기에 코드를 작성===

    ' 실행되는 PCB와 동일한 directory에 .hkp decrypt를 한 파일을 저장할 폴더 생성
    On Error Resume Next
    Dim folderHandler
    Set folderHandler = fso.CreateFolder(".\UpdatedData")
    On Error GoTo 0

    ' 툴 버전 확인 VX? (경로가 버전에 따라 달라지므로)
    ' Central Library 경로 넣기 (.lmc파일이 위치한곳) -> 여기서 PDB, Cel 정보얻음

    ' Decript 파일 경로 변수 정의
    Dim sToolVersion, sUtilityDir, sCentralLibDir, sPartDBDir, sCellDBDir, pcbPath, savedPath
    sToolVersion = "EEVX.2.11"
    sUtilityDir = "C:\MentorGraphics\" & sToolVersion & "\SDD_HOME\common\win64\bin\"

    ' Central Library의 경로를 넣음 (.lmc 파일이 위치한 곳)
    sCentralLibDir = "C:\WDIR\EEVX.2.11\3_Update_list_file_export\common\"
    sPartDBDir = sCentralLibDir & "\PartsDBLibs\"
    sCellDBDir = sCentralLibDir & "\CellDBLibs\"

    ' Decrypt된 데이터가 저장될 폴더
    pcbPath = pcbDoc.Path
    savedPath = pcbPath & "UpdatedData\"
    
    ' decript exe파일을 실행, PDB 데이터 파씽
    Win.run sUtilityDir & "PartsDB2HKP.exe -i " & sPartDBDir & "Discrete.pdb" & " -o " & savedPath & "decryptPDB.hkp -a"


    MsgBox("Work Well !")

    '===여기까지 작성===

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


