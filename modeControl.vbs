Call Scripting.AttachEvents(Application,"KeyIn")
Scripting.DontExit = True
Function KeyIn_OnProcessKeyin(KeyCom)
Com = Trim(KeyCom)
Exe = Split(Trim(KeyCom)," ")
If UBound(Exe) = -1 Then Exe = Array("")
Arg = Mid(KeyCom,Len(Exe(0)) + 1)
KeyIn_OnProcessKeyin = True
Select Case LCase(Exe(0))
  Case "qd"
    Gui.ActiveMode = epcbModeDrawing
  Case "qs"
    Gui.ActiveMode = epcbModeModeless
  Case "qp"
    Gui.ActiveMode = epcbModePlace
  Case "qr"
    Gui.ActiveMode = epcbModeRoute
  Case Else
    KeyIn_OnProcessKeyin = False
End Select
End Function