Attribute VB_Name = "BasToolBars"

Dim oApp As Outlook.Application
Dim oCB As Office.CommandBarButton
Dim oCBs As Office.CommandBars
Dim oMenuBar As Office.CommandBar
Dim oFolder As Outlook.MAPIFolder

Public Sub ToolbarCreate()
Dim x As String
x = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "LoadSMCToolbar")

If x = 0 Then
Exit Sub
End If

'Get the Application object for Outlook
Set oApp = Application

'Customize the Outlook Menu structure and toolbar
Set oCBs = oApp.ActiveExplorer.CommandBars
Set oMenuBar = oCBs.Add("SMC Toolbar", 1, False, True)
oMenuBar.Visible = True

A1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton1")
DoEvents
If A1 = "Yes" Then
Set obutton1 = oMenuBar.Controls.Add(msoControlButton, , , , True)
A2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Caption")
A3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Link")
A4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Icon")
DoEvents
obutton1.Caption = A2
obutton1.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton1.ToolTipText = A3
obutton1.FaceId = A4
obutton1.Style = msoButtonIconAndCaption
obutton1.Enabled = True
End If
DoEvents
DoEvents

B1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton2")
DoEvents
If B1 = "Yes" Then
Set obutton2 = oMenuBar.Controls.Add(msoControlButton, , , , True)
B2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button2Caption")
B3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button2Link")
B4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button2Icon")
DoEvents
obutton2.Caption = B2
obutton2.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton2.ToolTipText = B3
obutton2.FaceId = B4
obutton2.Style = msoButtonIconAndCaption
obutton2.Enabled = True
End If
DoEvents
DoEvents

C1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton3")
DoEvents
If C1 = "Yes" Then
Set obutton3 = oMenuBar.Controls.Add(msoControlButton, , , , True)
C2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button3Caption")
C3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button3Link")
C4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button3Icon")
DoEvents
obutton3.Caption = C2
obutton3.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton3.ToolTipText = C3
obutton3.FaceId = C4
obutton3.Style = msoButtonIconAndCaption
obutton3.Enabled = True
End If
DoEvents
DoEvents

D1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton4")
DoEvents
If D1 = "Yes" Then
Set obutton4 = oMenuBar.Controls.Add(msoControlButton, , , , True)
D2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button4Caption")
D3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button4Link")
D4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button4Icon")
DoEvents
obutton4.Caption = D2
obutton4.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton4.ToolTipText = D3
obutton4.FaceId = D4
obutton4.Style = msoButtonIconAndCaption
obutton4.Enabled = True
End If
DoEvents
DoEvents

E1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton5")
DoEvents
If E1 = "Yes" Then
Set obutton5 = oMenuBar.Controls.Add(msoControlButton, , , , True)
E2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button5Caption")
E3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button5Link")
E4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button5Icon")
DoEvents
obutton5.Caption = E2
obutton5.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton5.ToolTipText = E3
obutton5.FaceId = E4
obutton5.Style = msoButtonIconAndCaption
obutton5.Enabled = True
End If
DoEvents
DoEvents

F1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton6")
DoEvents
If F1 = "Yes" Then
Set obutton6 = oMenuBar.Controls.Add(msoControlButton, , , , True)
F2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button6Caption")
F3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button6Link")
F4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button6Icon")
DoEvents
obutton6.Caption = F2
obutton6.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton6.ToolTipText = F3
obutton6.FaceId = F4
obutton6.Style = msoButtonIconAndCaption
obutton6.Enabled = True
End If
DoEvents
DoEvents

G1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton7")
DoEvents
If G1 = "Yes" Then
Set obutton7 = oMenuBar.Controls.Add(msoControlButton, , , , True)
G2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button7Caption")
G3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button7Link")
G4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button7Icon")
DoEvents
obutton7.Caption = G2
obutton7.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton7.ToolTipText = G3
obutton7.FaceId = G4
obutton7.Style = msoButtonIconAndCaption
obutton7.Enabled = True
End If
DoEvents
DoEvents

H1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton8")
DoEvents
If H1 = "Yes" Then
Set obutton8 = oMenuBar.Controls.Add(msoControlButton, , , , True)
H2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button8Caption")
H3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button8Link")
H4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button8Icon")
DoEvents
obutton8.Caption = H2
obutton8.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton8.ToolTipText = H3
obutton8.FaceId = H4
obutton8.Style = msoButtonIconAndCaption
obutton8.Enabled = True
End If
DoEvents
DoEvents

I1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton9")
DoEvents
If I1 = "Yes" Then
Set obutton9 = oMenuBar.Controls.Add(msoControlButton, , , , True)
I2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button9Caption")
I3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button9Link")
I4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button9Icon")
DoEvents
obutton9.Caption = I2
obutton9.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton9.ToolTipText = I3
obutton9.FaceId = I4
obutton9.Style = msoButtonIconAndCaption
obutton9.Enabled = True
End If
DoEvents
DoEvents

J1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Createbutton0")
DoEvents
If J1 = "Yes" Then
Set obutton0 = oMenuBar.Controls.Add(msoControlButton, , , , True)
J2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button0Caption")
J3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button0Link")
J4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "button0Icon")
DoEvents
obutton0.Caption = J2
obutton0.HyperlinkType = msoCommandBarButtonHyperlinkOpen
obutton0.ToolTipText = J3
obutton0.FaceId = J4
obutton0.Style = msoButtonIconAndCaption
obutton0.Enabled = True
End If
DoEvents
DoEvents

Set obutton1 = Nothing
Set obutton2 = Nothing
Set obutton3 = Nothing
Set obutton4 = Nothing
Set obutton5 = Nothing
Set obutton6 = Nothing
Set obutton7 = Nothing
Set obutton8 = Nothing
Set obutton9 = Nothing
Set obutton0 = Nothing
Set oApp = Nothing
Set oCBs = Nothing
Set oMenuBar = Nothing
Set oNS = Nothing
Set oCB = Nothing
Set oFolder = Nothing
End Sub
