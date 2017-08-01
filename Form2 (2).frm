VERSION 5.00
Begin VB.Form login3_frm 
   Caption         =   "LogIn"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Palette         =   "Form2 (2).frx":0000
   Picture         =   "Form2 (2).frx":19932C
   ScaleHeight     =   10875
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   10875
      Left            =   14385
      Picture         =   "Form2 (2).frx":1B27D2
      ScaleHeight     =   10875
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton login_cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Login"
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   4095
   End
   Begin VB.CommandButton exit_cmd 
      BackColor       =   &H8000000D&
      Caption         =   "Exit"
      Height          =   315
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CheckBox show_pass_chk 
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   5160
      Width           =   230
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
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5520
      MaxLength       =   255
      TabIndex        =   3
      Text            =   "PASSWORD"
      Top             =   4560
      Width           =   3255
   End
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
      Height          =   405
      Left            =   5520
      MaxLength       =   255
      TabIndex        =   1
      Text            =   "USER NAME"
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton clr_cmd 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8760
      TabIndex        =   4
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton clr_cmd 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   2
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   1800
      Picture         =   "Form2 (2).frx":1C80B1
      Top             =   8040
      Width           =   4425
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   6240
      Picture         =   "Form2 (2).frx":1C951E
      Top             =   8040
      Width           =   3990
   End
   Begin VB.Image Image4 
      Height          =   900
      Left            =   10200
      Picture         =   "Form2 (2).frx":1CA721
      Top             =   8040
      Width           =   2370
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Password"
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
End
Attribute VB_Name = "login3_frm"
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

Private Sub exit_cmd_Click()
    Unload Me
End Sub

Private Sub login_cmd_Click()               ' for login button
    Set rs_login = New ADODB.Recordset
    rs_login.Open "select * from login_t where userid='" & usr_id_txt.Text & "' And  Pass='" & pass_txt.Text & "'", bms_cn, 3, 3
    
    
    
    If rs_login.RecordCount > 0 Then
        If ((TimeValue(Now()) >= rs_login.Fields("time1")) And (TimeValue(Now()) <= rs_login.Fields("time2"))) Or ((TimeValue(Now()) >= rs_login.Fields("time2")) And (TimeValue(Now()) <= rs_login.Fields("time1"))) Then
            bms_mdi.username = rs_login.Fields("username")
            bms_mdi.loginid = usr_id_txt.Text
            bms_mdi.user_name_cmd.Caption = rs_login.Fields("username")
            Unload Me
            bms_mdi.Show
        Else
            MsgBox "You can Not Login Now | You can login only between " + str(rs_login.Fields("time1")) + "-" + str(rs_login.Fields("time2")), vbCritical
        End If
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
