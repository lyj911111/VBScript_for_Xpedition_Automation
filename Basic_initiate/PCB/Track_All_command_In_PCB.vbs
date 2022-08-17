'
' PCB상에 존재하는 모든 command의 이름과 id를 출력
' 실행시, 현재 .pcb 파일 실행 경로에 "command_output.txt" 가 생성됨
' 그곳에 보면 모든 command의 이름과 id가 담겨있음
' 예) 
' &New... : 57600
' &Open... : 57601
'

' Create an output file
Set filesys = CreateObject("Scripting.FileSystemObject")
file = ".\command_output.txt"
Set filetxt = filesys.CreateTextFile(file, True)

' Get the document menu bar object
Dim docMenuBar
Set docMenuBar = Gui.CommandBars("Document Menu Bar")

' Walk through all menu in the menu bar
' and write out its name and command id
' to a file
xTab = vbTab
For i = 1 To docMenuBar.Controls.Count
  Set menu = docMenuBar.Controls.Item(i)
  filetxt.WriteLine "+" & menu.Caption
  Call WriteMenuIDs(menu)
  filetxt.WriteLine
Next

' Subroutine to write out menu item name
' and its command ID
Sub WriteMenuIDs(menu)
  Set menuCtrls = menu.Controls
  For j = 1 To menuCtrls.Count
    cmdName = menuCtrls.Item(j).Caption
    On Error Resume Next
    id = menuCtrls.Item(j).Id
    If Err Then
      ' CommandBarPopup doesn't support Id property
      Err.Clear
      filetxt.WriteLine xTab & cmdName
      saveTab = xTab
      xTab = xTab & vbTab
      Call WriteMenuIDs(menuCtrls.Item(j))
      xTab = saveTab
    ElseIf id <> 0 Then
      ' Don't write out separator whose command id is 0
      filetxt.WriteLine xTab & cmdName & " : " & id
    End If
  Next
End Sub

MsgBox "fini"
