1  Option Explicit
2  
3  ' Add any type libraries to be used. 
4  Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
5  
6  ' Get the Application object 
7  Dim pcbAppObj
8  Set pcbAppObj = Application
9  
10  ' Get the active document
11  Dim pcbDocObj
12  Set pcbDocObj = pcbAppObj.ActiveDocument
13  
14  ' License the document
15  ValidateServer(pcbDocObj)
16  
17  ' Get the vias collection
18  Dim viaColl
19  Set viaColl = pcbDocObj.Vias
20  
21  ' Get the number of vias in collection
22  Dim countInt
23  countInt = viaColl.Count
24  
25  MsgBox("There are " & countInt & " vias.")
26  
27  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
28  'Local functions
29  
30  ' Server validation function
31  Function ValidateServer(docObj)
32      
33      Dim keyInt
34      Dim licenseTokenInt
35      Dim licenseServerObj
36  
37      ' Ask Expedition’s document for the key
38      keyInt = docObj.Validate(0)
39  
40      ' Get license server
41      Set licenseServerObj = 
                     CreateObject("MGCPCBAutomationLicensing.Application")
42  
43      ' Ask the license server for the license token
44      licenseTokenInt = licenseServerObj.GetToken(keyInt)
45  
46      ' Release license server
47      Set licenseServerObj = nothing
48  
49      ' Turn off error messages (validate may fail if the 
       ' token is incorrect)
50      On Error Resume Next
51      Err.Clear
52  
53      ' Ask the document to validate the license token
54      docObj.Validate(licenseTokenInt)
55      If Err Then
56          ValidateServer = 0    
57      Else
58          ValidateServer = 1
59      End If
60
61  End Function