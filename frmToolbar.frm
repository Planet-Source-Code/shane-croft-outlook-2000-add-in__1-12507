VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmToolbar 
   Caption         =   "Customize SMC Toolbar"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "frmToolbar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Set Button"
      Height          =   315
      Left            =   4680
      TabIndex        =   29
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox TxtTo 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Text            =   "200"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox TxtFrom 
      Height          =   285
      Left            =   240
      TabIndex        =   23
      Text            =   "0"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enable Buttons"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox Check10 
         Caption         =   "10"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox Check9 
         Caption         =   "9"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox Check8 
         Caption         =   "8"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check7 
         Caption         =   "7"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check6 
         Caption         =   "6"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check5 
         Caption         =   "5"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox Check4 
         Caption         =   "4"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "3"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmToolbar.frx":0442
      Left            =   3240
      List            =   "frmToolbar.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Button Icons"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "To"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   $"frmToolbar.frx":0446
      Height          =   1215
      Left            =   2040
      TabIndex        =   22
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Please Enter A Number. Example: 125"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "What Icon Would You Like For This Button?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "What Would You Like To Call This Button?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "What Would You Like The Button To Open?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Choose a Button"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   3960
      Width           =   3975
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oApp As Outlook.Application
Dim oCB As Office.CommandBarButton
Dim oCBs As Office.CommandBars
Dim oMenuBar As Office.CommandBar
Dim oFolder As Outlook.MAPIFolder
Public Sub LoadCheckBox()
On Error Resume Next
A1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton1")
If A1 = "Yes" Then
Check1.Value = 1
DoEvents
End If

A2 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton2")
If A2 = "Yes" Then
Check2.Value = 1
DoEvents
End If

A3 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton3")
If A3 = "Yes" Then
Check3.Value = 1
DoEvents
End If

A4 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton4")
If A4 = "Yes" Then
Check4.Value = 1
DoEvents
End If

A5 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton5")
If A5 = "Yes" Then
Check5.Value = 1
DoEvents
End If

A6 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton6")
If A6 = "Yes" Then
Check6.Value = 1
DoEvents
End If

A7 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton7")
If A7 = "Yes" Then
Check7.Value = 1
DoEvents
End If

A8 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton8")
If A8 = "Yes" Then
Check8.Value = 1
DoEvents
End If

A9 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton9")
If A9 = "Yes" Then
Check9.Value = 1
DoEvents
End If

A0 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton0")
If A0 = "Yes" Then
Check10.Value = 1
DoEvents
End If
End Sub
Public Sub LoadComboBox()
On Error Resume Next
If Check1.Value = 1 Then
Combo1.AddItem "1"
DoEvents
End If

If Check2.Value = 1 Then
Combo1.AddItem "2"
DoEvents
End If

If Check3.Value = 1 Then
Combo1.AddItem "3"
DoEvents
End If

If Check4.Value = 1 Then
Combo1.AddItem "4"
DoEvents
End If

If Check5.Value = 1 Then
Combo1.AddItem "5"
DoEvents
End If

If Check6.Value = 1 Then
Combo1.AddItem "6"
DoEvents
End If

If Check7.Value = 1 Then
Combo1.AddItem "7"
DoEvents
End If

If Check8.Value = 1 Then
Combo1.AddItem "8"
DoEvents
End If

If Check9.Value = 1 Then
Combo1.AddItem "9"
DoEvents
End If

If Check10.Value = 1 Then
Combo1.AddItem "0"
DoEvents
End If
End Sub
Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 0 Then
Combo1.Clear
DoEvents
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton1", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "1"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton1", "Yes"
End If
End Sub

Private Sub Check10_Click()
On Error Resume Next
If Check10.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton0", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "10"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton0", "Yes"
End If
End Sub

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton2", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "2"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton2", "Yes"
End If
End Sub

Private Sub Check3_Click()
On Error Resume Next
If Check3.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton3", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "3"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton3", "Yes"
End If
End Sub

Private Sub Check4_Click()
On Error Resume Next
If Check4.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton4", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "4"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton4", "Yes"
End If
End Sub

Private Sub Check5_Click()
On Error Resume Next
If Check5.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton5", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "5"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton5", "Yes"
End If
End Sub

Private Sub Check6_Click()
On Error Resume Next
If Check6.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton6", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "6"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton6", "Yes"
End If
End Sub

Private Sub Check7_Click()
On Error Resume Next
If Check7.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton7", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "7"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton7", "Yes"
End If
End Sub

Private Sub Check8_Click()
On Error Resume Next
If Check8.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton8", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "8"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton8", "Yes"
End If
End Sub

Private Sub Check9_Click()
On Error Resume Next
If Check9.Value = 0 Then
Combo1.Clear
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton9", "No"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Caption", "-"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Link", "\"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Icon", "0"
DoEvents
DoEvents
DoEvents
Call LoadComboBox
Else
Combo1.AddItem "9"
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "CreateButton9", "Yes"
End If
End Sub

Private Sub Combo1_Click()
On Error Resume Next

If Combo1.Text = "1" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Icon")
Exit Sub
End If

If Combo1.Text = "2" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Icon")
Exit Sub
End If

If Combo1.Text = "3" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Icon")
Exit Sub
End If

If Combo1.Text = "4" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Icon")
Exit Sub
End If

If Combo1.Text = "5" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Icon")
Exit Sub
End If

If Combo1.Text = "6" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Icon")
Exit Sub
End If

If Combo1.Text = "7" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Icon")
Exit Sub
End If

If Combo1.Text = "8" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Icon")
Exit Sub
End If

If Combo1.Text = "9" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Icon")
Exit Sub
End If

If Combo1.Text = "10" Then
Text1.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Caption")
Text2.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Link")
Text3.Text = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Icon")
Exit Sub
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
MsgBox "You will need to restart Outlook 2000 for changes to take effect."
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim x As Long
Dim x2 As Long
x = TxtFrom.Text
x2 = TxtTo.Text
Set oApp = Application
'Customize the Outlook Menu structure and toolbar
oApp.ActiveExplorer.CommandBars("Button Icons").Delete
Set oCBs = oApp.ActiveExplorer.CommandBars
Set oMenuBar = oCBs.Add("Button Icons", , False, True)
oMenuBar.Visible = False

Do Until x = x2 + 1
Label1.Caption = "Getting Button Icon # " & x
DoEvents
Set obutton1 = oMenuBar.Controls.Add(msoControlButton, , , , True)
obutton1.Caption = x
obutton1.Style = msoButtonIcon
obutton1.FaceId = x
obutton1.Enabled = True
x = x + 1
Set obutton1 = Nothing
Loop

oMenuBar.Width = 600
oMenuBar.Visible = True

Set obutton1 = Nothing
Set oApp = Nothing
Set oCBs = Nothing
Set oMenuBar = Nothing
Set oNS = Nothing
Set oCB = Nothing
Set oFolder = Nothing

End Sub

Private Sub Command4_Click()
On Error Resume Next
CD1.CancelError = False
CD1.FileName = ""
CD1.Flags = cdlOFNFileMustExist
CD1.Filter = "All Files (*.*)|*.*"
CD1.ShowOpen
Text2.Text = CD1.FileName
End Sub

Private Sub Command5_Click()
On Error Resume Next

If Text1.Text = "" Then
MsgBox "You must enter a name for the button."
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "You must enter a file path for the button."
Exit Sub
End If

If Text3.Text = "" Then
MsgBox "You must enter a number for the button icon."
Exit Sub
End If

If Combo1.Text = "1" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button1Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "2" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button2Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "3" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button3Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "4" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button4Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "5" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button5Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "6" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button6Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "7" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button7Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "8" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button8Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "9" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button9Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

If Combo1.Text = "10" Then
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Caption", Text1.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Link", Text2.Text
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "Button0Icon", Text3.Text
DoEvents
DoEvents
Combo1.Text = ""
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
Call LoadCheckBox
DoEvents
DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Form1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.SetFocus
End Sub

Private Sub Timer1_Timer()
If Combo1.Text = "" Then
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Command5.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Command5.Enabled = True
End If
End Sub

