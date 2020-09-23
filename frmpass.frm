VERSION 5.00
Begin VB.Form frmpass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please Enter Password To Enter SMC Settings"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   Icon            =   "frmpass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2288
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   848
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Type Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function DeEncrypt(What As String) As String
    Dim Before$, After$, EpN%, Dracula%, Aeneima$, DeMoNs$
    Before$ = " ¿¡@#$%^&*()_+|01²³456789ÀbÁdÂÃghÄjklmÅÒÓqÔÕÖÙvwÛÜz.-~,AàáâãFGHäJKåMNØ¶QR§TÚVWX¥Z?!23acefinoprstuxyBCDEILOPSUY"
    After$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂÃÄÅÒÓÔÕÖÙÛÜàáâãäåØ¶§Ú¥"
    For EpN% = 1 To Len(What)
        Dracula% = InStr(Before$, Mid(What, EpN%, 1))
        If Not Dracula% = 0 Then
            Aeneima$ = Mid(After$, Dracula%, 1)
            DeMoNs$ = DeMoNs$ + Aeneima$
        End If
    Next
    DeEncrypt = DeMoNs$
End Function

Private Sub Command1_Click()
If Text1.Text = Text2.Text Then
Form1.Show
Unload Me
Else
MsgBox "Wrong password. Please try again"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Z1 = regQuery_A_Key(HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "PWOn")
DoEvents
If Z1 = "Yes" Then
Dim intfile As Integer
Dim pass As String
intfile = FreeFile
  Open App.Path & "\Settings.ini" For Input As #intfile
  Input #intfile, pass
  Text3.Text = pass
  Close #intfile
DoEvents
DoEvents
Text2.Text = DeEncrypt(Text3.Text)
Exit Sub
End If
Form1.Show
Unload Me
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If

End Sub
