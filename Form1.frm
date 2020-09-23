VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SMC Outlook 2000 Settings & Options"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Set Settings Password"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Customize SMC Toolbar"
      Height          =   375
      Left            =   585
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Load Toolbar At Outlook Startup"
      Height          =   375
      Left            =   285
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
On Error Resume Next
regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "LoadSMCToolbar", Check1.Value
DoEvents
DoEvents
DoEvents
End
End Sub

Private Sub Command3_Click()
frmsetpass.Show
End Sub

Private Sub Command4_Click()
frmToolbar.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Check1.Value = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "LoadSMCToolbar")
End Sub
