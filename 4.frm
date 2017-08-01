VERSION 5.00
Begin VB.Form login1_frm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LogIn"
   ClientHeight    =   10875
   ClientLeft      =   -2895
   ClientTop       =   3765
   ClientWidth     =   13545
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   10875
      Left            =   13410
      ScaleHeight     =   10875
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   7680
      TabIndex        =   1
      Top             =   2640
      Width           =   5175
      Begin VB.TextBox usr_id_txt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   960
         MaxLength       =   255
         TabIndex        =   6
         Text            =   "USER NAME"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox pass_txt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   255
         TabIndex        =   5
         Text            =   "PASSWORD"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CheckBox show_pass_chk 
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   2640
         Width           =   230
      End
      Begin VB.CommandButton exit_cmd 
         BackColor       =   &H8000000D&
         Caption         =   "Exit"
         Height          =   555
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton login_cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Login"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   1350
         Left            =   1920
         Top             =   0
         Width           =   1350
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   480
         Top             =   1920
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Password"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
   End
End
Attribute VB_Name = "login1_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_login As ADODB.Recordset






Private Sub clr_cmd_Click(Index As Integer)
    If Index = 0 Then
      usr_id_txt.Text = ""
    Else
      pass_txt.Text = ""
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub exit_cmd_Click()
    Unload Me
End Sub

Private Sub login_cmd_Click()               ' for login button
    Set rs_login = New ADODB.Recordset
    rs_login.Open "select * from login_t where userid='" & usr_id_txt.Text & "' And  Pass='" & pass_txt.Text & "'", bms_cn, 3, 3
    
    If rs_login.RecordCount > 0 Then
        Unload Me
        bms_mdi.Show
    Else
        MsgBox "worng UserId/Password", vbCritical
    End If
        
End Sub



Private Sub pass_txt_GotFocus()
    If pass_txt.Text = "PASSWORD" Then
        pass_txt.Text = ""
        If show_pass_chk.value = 0 Then
            pass_txt.PasswordChar = "*"
        End If
    End If
End Sub

Private Sub pass_txt_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then
           Call login_cmd_Click
         End If
End Sub

Private Sub pass_txt_LostFocus()
        If pass_txt.Text = "" Then
           If show_pass_chk.value = 0 Then
              pass_txt.PasswordChar = ""
              End If
           pass_txt.Text = "PASSWORD"
        End If
End Sub

Private Sub show_pass_chk_Click()       ' For show Hide Password Char
     If show_pass_chk.value = 1 Then
        pass_txt.PasswordChar = ""
    Else
        pass_txt.PasswordChar = "*"
    End If
End Sub








Private Sub usr_id_txt_GotFocus()
    If usr_id_txt.Text = "USER NAME" Then
         usr_id_txt.Text = ""
    End If
End Sub

Private Sub usr_id_txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
           pass_txt.SetFocus
    End If
End Sub

Private Sub usr_id_txt_LostFocus()
     If usr_id_txt.Text = "" Then
         usr_id_txt.Text = "USER NAME"
    End If
End Sub
