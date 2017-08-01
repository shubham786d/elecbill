VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form createuser_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "user form"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "createuser_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7440
      TabIndex        =   21
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "HH.mm"
      Format          =   187826179
      CurrentDate     =   42637
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   6960
      TabIndex        =   17
      Top             =   3360
      Width           =   3495
      Begin VB.OptionButton normalusr_opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Normal User"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton admin_opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin"
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.OptionButton bothrights_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Both"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10320
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Tranrights_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction Tab"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8760
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton masterright_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Master Tab"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7440
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   10800
      Picture         =   "createuser_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton seepass_cmd 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9720
      TabIndex        =   11
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox password_txt 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox userId_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   6
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox username_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      Picture         =   "createuser_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      Picture         =   "createuser_frm.frx":15B3D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   5040
      Picture         =   "createuser_frm.frx":1610A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   9240
      Picture         =   "createuser_frm.frx":16637
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   9240
      TabIndex        =   22
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "HH.mm"
      Format          =   187826179
      CurrentDate     =   42637
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login Time Between:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Rights :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   3480
      Top             =   360
      Width           =   8295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   5895
      Left            =   3480
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "createuser_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As ADODB.Recordset
Public state As Integer ' 1:insert 2:update
Public oldname As String

Private Sub admin_opt_Click()
    Label5.Visible = False
    masterright_opt.Visible = False
    Tranrights_opt.Visible = False
    bothrights_opt.Visible = False
End Sub

Private Sub new_cmd_Click()
    username_txt.Enabled = True
    userId_txt.Enabled = True
    password_txt.Enabled = True
    seepass_cmd.Enabled = True
    admin_opt.Enabled = True
    normalusr_opt.Enabled = True
    masterright_opt.Enabled = True
    Tranrights_opt.Enabled = True
    bothrights_opt.Enabled = True
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    s_cmd.Enabled = True
    
    admin_opt = True
    state = 1
End Sub

Public Sub normalusr_opt_Click()
    
   
    
    Label5.Visible = True
    masterright_opt.Visible = True
    Tranrights_opt.Visible = True
    bothrights_opt.Visible = True
    masterright_opt = True
End Sub

Private Sub s_cmd_Click()
    Select Case state
        Case 1
            If username_txt.Text <> "" Then
                If userId_txt.Text <> "" Then
                    If password_txt.Text <> "" Then
        
                            Set rst = New ADODB.Recordset
                            rst.CursorLocation = adUseClient
                               
                            rst.Open ("select * from login_t where userid='" & userId_txt.Text & "'"), bms_cn, 3, 3
                            If rst.RecordCount = 0 Then
                                Dim rights As Integer
                                Dim utype As String
                                
                                If (admin_opt.value = True) Then
                                    rights = 4
                                    utype = "admin"
                                ElseIf (bothrights_opt.value = True) Then
                                    rights = 3
                                    utype = "normal user"
                                ElseIf (Tranrights_opt.value = True) Then
                                    rights = 2
                                    utype = "normal user"
                                ElseIf (masterright_opt.value = True) Then
                                    rights = 1
                                    utype = "normal user"
                                End If
                                
                                
                                Dim str As String
                                str = "insert into login_t values('" & username_txt.Text & "','" & userId_txt.Text & "','" & password_txt.Text & "','" & rights & "','" & utype & "','','" & TimeValue(DTPicker1.value) & "','" & TimeValue(DTPicker2.value) & "')"
                                insert (str)
                                state = 3
                                
                                username_txt.Enabled = False
                                userId_txt.Enabled = False
                                password_txt.Enabled = False
                                seepass_cmd.Enabled = False
                                admin_opt.Enabled = False
                                normalusr_opt.Enabled = False
                                masterright_opt.Enabled = False
                                Tranrights_opt.Enabled = False
                                bothrights_opt.Enabled = False
                                s_cmd.Enabled = False
                                del_cmd.Enabled = False
                                
                                username_txt = ""
                                userId_txt = ""
                                password_txt = ""
                                
                                Label5.Visible = False
                                masterright_opt.Visible = False
                                Tranrights_opt.Visible = False
                                bothrights_opt.Visible = False
                                masterright_opt = False
                                DTPicker1.Enabled = False
                                DTPicker2.Enabled = False
                                admin_opt = True
                                
                                MsgBox "New Record Saved Successfully", vbInformation
                            
                            Else
                                MsgBox "user already exist with this User id Please input another User Id", vbInformation
                            End If
        
                    Else
                        MsgBox "Please Input User's Password", vbInformation
                    End If
                Else
                    MsgBox "Please Input User's ID", vbInformation
                End If
            Else
                MsgBox "Please Input User's Name", vbInformation
            End If
        Case 2
            If username_txt <> "" Then
                If userId_txt.Text <> "" Then
                    If password_txt.Text <> "" Then
                        If userId_txt.Text <> oldname Then
                            Set rst = New ADODB.Recordset
                            rst.CursorLocation = adUseClient
                               
                            rst.Open ("select * from login_t where userid='" & userId_txt.Text & "'"), bms_cn, 3, 3
                            If rst.RecordCount <> 0 Then
                                MsgBox "user already exist with this User id Please input another User Id", vbInformation
                                Exit Sub
                            End If
                        End If
                        
                        
                        
                            If (admin_opt.value = True) Then
                                    rights = 4
                                    utype = "admin"
                                ElseIf (bothrights_opt.value = True) Then
                                    rights = 3
                                    utype = "normal user"
                                ElseIf (Tranrights_opt.value = True) Then
                                    rights = 2
                                    utype = "normal user"
                                ElseIf (masterright_opt.value = True) Then
                                    rights = 1
                                    utype = "normal user"
                            End If
                            
                            
                            str = "update login_t set username='" & username_txt.Text & "',userid='" & userId_txt.Text & "',pass='" & password_txt.Text & "',rights='" & rights & "',type='" & utype & "',time1='" & TimeValue(DTPicker1.value) & "',time2='" & TimeValue(DTPicker2.value) & "' where userid='" & oldname & "' "
                            update (str)
                            MsgBox "Record Updateted Successfully", vbInformation
                            state = 3
                                
                            username_txt.Enabled = False
                            userId_txt.Enabled = False
                            password_txt.Enabled = False
                            seepass_cmd.Enabled = False
                            admin_opt.Enabled = False
                            normalusr_opt.Enabled = False
                            masterright_opt.Enabled = False
                            Tranrights_opt.Enabled = False
                            bothrights_opt.Enabled = False
                            s_cmd.Enabled = False
                            del_cmd.Enabled = False
                            DTPicker1.Enabled = False
                            DTPicker2.Enabled = False
                            username_txt = ""
                            userId_txt = ""
                            password_txt = ""
                            
                            Label5.Visible = False
                            masterright_opt.Visible = False
                            Tranrights_opt.Visible = False
                            bothrights_opt.Visible = False
                            masterright_opt = False
                            
                            admin_opt = True
                       
                    Else
                        MsgBox "Please Input User's Password", vbInformation
                    End If
                Else
                    MsgBox "Please Input User's ID", vbInformation
                End If
            Else
            MsgBox "Please Input User's Name", vbInformation
        End If
    End Select
End Sub



Private Sub seepass_cmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
password_txt.PasswordChar = ""
End Sub

Private Sub seepass_cmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
password_txt.PasswordChar = "*"
End Sub

Private Sub src_cmd_Click()
    createusersrc_frm.Show vbModal
End Sub
