Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

' current .pcb data : pcbDoc ( e.g. Coporate.pcb )
' current Software Application : pcbApp ( e.g. Xpedition Layout )
Msgbox pcbDoc & " " & pcbApp