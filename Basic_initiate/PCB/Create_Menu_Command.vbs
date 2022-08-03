' 메뉴만들기

Dim docMenuBar
Set docMenuBar = Gui.CommandBars("Document Menu Bar")

Dim myMenu
Set myMenu = docMenuBar.Controls.Add(cmdControlPopup,,,-1)
' 내 메뉴 이름
myMenu.Caption = "test999"


' myMenu에 기능추가
Set myCntrls = myMenu.Controls
Set cmd = myCntrls.Add

' 추가된 메뉴의 text이름
cmd.Caption = "Zoom All"   
' 해당 text를 눌렀을 때 하는 Action 명령어
cmd.OnAction = "za"   

Set cmd = myCntrls.Add
' Zoom board를 눌렀을 때 zb 커맨드를 실행함
cmd.Caption = "Zoom board"   
cmd.OnAction = "zb"   

Set cmd = myCntrls.Add
' my vbs를 눌렀을 때 해당 경로에 있는 .vbs 스크립트가 실행됨
cmd.Caption = "my vbs"   
cmd.OnAction = "run myScript.vbs"  ' OnAction Property for a script

' Seperation 구분자 넣기
myCntrls.Add(cmdControlButtonSeparator)

' 하위 팝업창 메뉴 만들기
Dim mySubMenu
Set mySubMenu = myCntrls.Add(cmdControlPopup,,,-1)
mySubMenu.Caption = "Sub Menu"

' 하위메뉴에 명령문 넣기
Dim mySubMenuCntrls
Set mySubMenuCntrls = mySubMenu.Controls

' 첫번째 하위 메뉴
Set cmdSubMenu1 = mySubMenuCntrls.Add
cmdSubMenu1.Caption = "Setup Parameters"
cmdSubMenu1.ExecuteMethod = "setupParam"
cmdSubMenu1.Target = ScriptEngine
Scripting.DontExit = True

' 두번째 하위 메뉴
Set cmdSubMenu2 = mySubMenuCntrls.Add
cmdSubMenu2.Caption = "Stackup Editor"
cmdSubMenu2.ExecuteMethod = "StackupEditor"
cmdSubMenu2.Target = ScriptEngine
Scripting.DontExit = True

' 세번째 하위 메뉴
Set cmdSubMenu3 = mySubMenuCntrls.Add
cmdSubMenu3.Caption = "Constraint Manager"
cmdSubMenu3.ExecuteMethod = "ConstManager"
cmdSubMenu3.Target = ScriptEngine
Scripting.DontExit = True

Sub setupParam(nID)
    Gui.ProcessCommand "Setup->Setup Parameters", True
End Sub

Sub StackupEditor(nID)
    Gui.ProcessCommand "Setup->Stackup Editor", True
End Sub

Sub ConstManager(nID)
    Gui.ProcessCommand "Setup->Constraint Manager", True
End Sub
