VERSION 5.00
Begin VB.Form frmsetpass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Password For SMC Settings"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   Icon            =   "frmsetpass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2235
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox TxtPassword 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   915
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox TxtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox TxtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "To disable the password feature just make the password blank."
      Height          =   855
      Left            =   1395
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Retype Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
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
Attribute VB_Name = "frmsetpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Encrypt(What As String) As String
    Dim Before$, After$, EpN%, Dracula%, Aeneima$, DeMoNs$
    Before$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂÃÄÅÒÓÔÕÖÙÛÜàáâãäåØ¶§Ú¥"
    After$ = " ¿¡@#$%^&*()_+|01²³456789ÀbÁdÂÃghÄjklmÅÒÓqÔÕÖÙvwÛÜz.-~,AàáâãFGHäJKåMNØ¶QR§TÚVWX¥Z?!23acefinoprstuxyBCDEILOPSUY"
    For EpN% = 1 To Len(What)
        Dracula% = InStr(Before$, Mid(What, EpN%, 1))
        If Not Dracula% = 0 Then
            Aeneima$ = Mid(After$, Dracula%, 1)
            DeMoNs$ = DeMoNs$ + Aeneima$
        End If
    Next
    Encrypt = DeMoNs$
End Function
Private Sub Command1_Click()
On Error Resume Next

If TxtPass.Text = "" Then
    regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "PWOn", "No"
    DoEvents
    MsgBox "Password has been disabled."
    Unload Me
    Exit Sub
End If
    
If TxtPass.Text = TxtConfirm.Text Then
TxtPassword.Text = Encrypt(TxtConfirm.Text)
DoEvents
DoEvents
DoEvents
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TxtPassword.Text
    Close
    regCreate_Key_Value HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\SMC_Outlook_2000.Connect", "PWOn", "Yes"
    MsgBox "Password has been set."
    Unload Me
Else
MsgBox "Sorry you mistyped your password please try again."
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Form1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.SetFocus
End Sub

Private Sub TxtConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If
End Sub
