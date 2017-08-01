VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm pms_mdi 
   BackColor       =   &H8000000C&
   Caption         =   "Placement Management System"
   ClientHeight    =   9930
   ClientLeft      =   -5535
   ClientTop       =   1005
   ClientWidth     =   15120
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   15120
      TabIndex        =   3
      Top             =   1935
      Width           =   15120
      Begin VB.CommandButton alltabc_cmd 
         BackColor       =   &H0080C0FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   255
      End
      Begin TabDlg.SSTab frmtab 
         Height          =   6255
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   19320
         _ExtentX        =   34078
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   20
         Tab             =   6
         TabsPerRow      =   10
         TabHeight       =   794
         TabMaxWidth     =   3351
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "MDIForm1.frx":D4DE
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "c_tab_cmd(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "MDIForm1.frx":D4FA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "c_tab_cmd(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "MDIForm1.frx":D516
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "c_tab_cmd(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Tab 3"
         TabPicture(3)   =   "MDIForm1.frx":D532
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "c_tab_cmd(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Tab 4"
         TabPicture(4)   =   "MDIForm1.frx":D54E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         TabCaption(5)   =   "Tab 5"
         TabPicture(5)   =   "MDIForm1.frx":D56A
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         TabCaption(6)   =   "Tab 6"
         TabPicture(6)   =   "MDIForm1.frx":D586
         Tab(6).ControlEnabled=   -1  'True
         Tab(6).ControlCount=   0
         TabCaption(7)   =   "Tab 7"
         TabPicture(7)   =   "MDIForm1.frx":D5A2
         Tab(7).ControlEnabled=   0   'False
         Tab(7).ControlCount=   0
         TabCaption(8)   =   "Tab 8"
         TabPicture(8)   =   "MDIForm1.frx":D5BE
         Tab(8).ControlEnabled=   0   'False
         Tab(8).ControlCount=   0
         TabCaption(9)   =   "Tab 9"
         TabPicture(9)   =   "MDIForm1.frx":D5DA
         Tab(9).ControlEnabled=   0   'False
         Tab(9).ControlCount=   0
         TabCaption(10)  =   "Tab 10"
         TabPicture(10)  =   "MDIForm1.frx":D5F6
         Tab(10).ControlEnabled=   0   'False
         Tab(10).ControlCount=   0
         TabCaption(11)  =   "Tab 11"
         TabPicture(11)  =   "MDIForm1.frx":D612
         Tab(11).ControlEnabled=   0   'False
         Tab(11).ControlCount=   0
         TabCaption(12)  =   "Tab 12"
         TabPicture(12)  =   "MDIForm1.frx":D62E
         Tab(12).ControlEnabled=   0   'False
         Tab(12).ControlCount=   0
         TabCaption(13)  =   "Tab 13"
         TabPicture(13)  =   "MDIForm1.frx":D64A
         Tab(13).ControlEnabled=   0   'False
         Tab(13).ControlCount=   0
         TabCaption(14)  =   "Tab 14"
         TabPicture(14)  =   "MDIForm1.frx":D666
         Tab(14).ControlEnabled=   0   'False
         Tab(14).ControlCount=   0
         TabCaption(15)  =   "Tab 15"
         TabPicture(15)  =   "MDIForm1.frx":D682
         Tab(15).ControlEnabled=   0   'False
         Tab(15).ControlCount=   0
         TabCaption(16)  =   "Tab 16"
         TabPicture(16)  =   "MDIForm1.frx":D69E
         Tab(16).ControlEnabled=   0   'False
         Tab(16).ControlCount=   0
         TabCaption(17)  =   "Tab 17"
         TabPicture(17)  =   "MDIForm1.frx":D6BA
         Tab(17).ControlEnabled=   0   'False
         Tab(17).ControlCount=   0
         TabCaption(18)  =   "Tab 18"
         TabPicture(18)  =   "MDIForm1.frx":D6D6
         Tab(18).ControlEnabled=   0   'False
         Tab(18).ControlCount=   0
         TabCaption(19)  =   "Tab 19"
         TabPicture(19)  =   "MDIForm1.frx":D6F2
         Tab(19).ControlEnabled=   0   'False
         Tab(19).ControlCount=   0
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   3
            Left            =   -67560
            MaskColor       =   &H00FF8080&
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   2
            Left            =   -69480
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   1
            Left            =   -71400
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   0
            Left            =   -73320
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   0
      Picture         =   "MDIForm1.frx":D70E
      ScaleHeight     =   1875
      ScaleWidth      =   15060
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   12720
         Picture         =   "MDIForm1.frx":1002D
         ScaleHeight     =   375
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   0
         Width           =   495
         Begin VB.PictureBox Picture5 
            Height          =   15
            Left            =   0
            ScaleHeight     =   15
            ScaleWidth      =   495
            TabIndex        =   19
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "ADMIN"
         Height          =   400
         Left            =   13200
         Picture         =   "MDIForm1.frx":1045D
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   2055
      End
      Begin VB.PictureBox usr_win_pb 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   12720
         Picture         =   "MDIForm1.frx":10A67
         ScaleHeight     =   2415
         ScaleWidth      =   2535
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton Command6 
            Caption         =   "Command2"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Command2"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Command2"
            Height          =   375
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command2"
            Height          =   375
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   2175
         End
      End
      Begin TabDlg.SSTab menu 
         Height          =   1935
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   53535
         _ExtentX        =   94430
         _ExtentY        =   3413
         _Version        =   393216
         Tab             =   1
         TabHeight       =   706
         TabMaxWidth     =   5292
         WordWrap        =   0   'False
         BackColor       =   16777215
         ForeColor       =   16777215
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "MDIForm1.frx":11542
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "plc_ro_mst_mnu_fr"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "std_mst_mnu_fr"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "pd_mst_mnu_fr"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "ro_typ_mst_mnu_fr"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "co_mst_mnu_fr"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Shape1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "MDIForm1.frx":11EBD
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Shape2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Shape5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame2"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame3"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "MDIForm1.frx":12C89
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Shape4"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   1450
            Left            =   6120
            TabIndex        =   39
            Top             =   480
            Width           =   1575
            Begin VB.Image Image2 
               Height          =   1050
               Left            =   240
               Picture         =   "MDIForm1.frx":136AD
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1450
            Left            =   4080
            TabIndex        =   34
            Top             =   480
            Width           =   1575
            Begin VB.PictureBox Picture6 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1050
               Left            =   240
               Picture         =   "MDIForm1.frx":140B7
               ScaleHeight     =   1050
               ScaleWidth      =   1050
               TabIndex        =   35
               Top             =   240
               Width           =   1050
               Begin VB.Label Label8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "    Course Master"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   36
                  Top             =   1320
                  Width           =   1455
               End
            End
            Begin VB.Label Label10 
               Caption         =   "    Course Master"
               Height          =   15
               Left            =   120
               TabIndex        =   38
               Top             =   1060
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "Course Master"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   1200
               Width           =   1095
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1450
            Left            =   1080
            TabIndex        =   27
            Top             =   480
            Width           =   1575
            Begin VB.PictureBox Picture7 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   900
               Left            =   360
               Picture         =   "MDIForm1.frx":14B08
               ScaleHeight     =   900
               ScaleWidth      =   900
               TabIndex        =   40
               Top             =   240
               Width           =   900
               Begin VB.Label Label12 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "    Course Master"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   41
                  Top             =   1320
                  Width           =   1455
               End
            End
            Begin VB.Label Label7 
               Caption         =   "Student Master"
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   1200
               Width           =   1095
            End
         End
         Begin VB.Frame plc_ro_mst_mnu_fr 
            Height          =   1450
            Left            =   -65880
            TabIndex        =   25
            Top             =   420
            Width           =   1575
            Begin VB.Label Label6 
               Caption         =   "Placement Round            Master"
               Height          =   375
               Left            =   120
               TabIndex        =   26
               Top             =   960
               Width           =   1380
            End
            Begin VB.Image plc_ro_mst_mnu_img 
               Height          =   1050
               Left            =   240
               Picture         =   "MDIForm1.frx":15337
               Top             =   120
               Width           =   1050
            End
         End
         Begin VB.Frame std_mst_mnu_fr 
            Height          =   1450
            Left            =   -73680
            TabIndex        =   15
            Top             =   420
            Width           =   1575
            Begin VB.Label Label2 
               Caption         =   "Student Master"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Image std_mst_mnu_img 
               Height          =   1050
               Left            =   240
               Picture         =   "MDIForm1.frx":15F44
               Top             =   120
               Width           =   1050
            End
         End
         Begin VB.Frame pd_mst_mnu_fr 
            Height          =   1450
            Left            =   -69720
            TabIndex        =   14
            Top             =   420
            Width           =   1575
            Begin VB.Label Label5 
               Caption         =   "Placement Drive            Master"
               Height          =   440
               Left            =   120
               TabIndex        =   22
               Top             =   960
               Width           =   1335
            End
            Begin VB.Image pd_mst_mnu_img 
               Height          =   1050
               Left            =   240
               Picture         =   "MDIForm1.frx":19BAE
               Top             =   120
               Width           =   930
            End
         End
         Begin VB.Frame ro_typ_mst_mnu_fr 
            Height          =   1450
            Left            =   -67800
            TabIndex        =   21
            Top             =   420
            Width           =   1575
            Begin VB.Label Label3 
               Caption         =   "Round Type Master"
               Height          =   255
               Left            =   45
               TabIndex        =   23
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Image ro_typ_mst_mnu_img 
               Height          =   1125
               Left            =   80
               Picture         =   "MDIForm1.frx":1A492
               Top             =   120
               Width           =   1380
            End
         End
         Begin VB.Frame co_mst_mnu_fr 
            Height          =   1450
            Left            =   -71640
            TabIndex        =   10
            Top             =   420
            Width           =   1575
            Begin VB.PictureBox co_mst_mnu_img 
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   945
               Left            =   240
               Picture         =   "MDIForm1.frx":1CDCB
               ScaleHeight     =   945
               ScaleWidth      =   1050
               TabIndex        =   11
               Top             =   120
               Width           =   1050
               Begin VB.Label Label1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "    Course Master"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   12
                  Top             =   1320
                  Width           =   1455
               End
            End
            Begin VB.Label Label4 
               Caption         =   "Course Master"
               Height          =   255
               Left            =   240
               TabIndex        =   24
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label co_mst_mnu_lbl 
               Caption         =   "    Course Master"
               Height          =   15
               Left            =   120
               TabIndex        =   13
               Top             =   1060
               Width           =   1335
            End
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   45
            Left            =   3120
            Top             =   0
            Width           =   2895
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   45
            Left            =   -68880
            Top             =   0
            Width           =   3015
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   45
            Left            =   3360
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   50
            Left            =   -75000
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   10455
         TabIndex        =   1
         Top             =   0
         Width           =   10455
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000D&
         FillColor       =   &H80000005&
         Height          =   45
         Left            =   3120
         Top             =   1320
         Width           =   3015
      End
   End
   Begin VB.Menu pms_mst 
      Caption         =   "&Master"
      Begin VB.Menu mst_std 
         Caption         =   "&Student Master"
         Shortcut        =   ^S
      End
      Begin VB.Menu mst_comp 
         Caption         =   "&Company Master"
      End
      Begin VB.Menu mst_pd 
         Caption         =   "&Placement Drive Master"
         Shortcut        =   ^P
      End
      Begin VB.Menu mst_co 
         Caption         =   "&Course Master"
         Shortcut        =   ^C
      End
      Begin VB.Menu mst_ro 
         Caption         =   "&Rounds Master"
         Shortcut        =   ^R
      End
      Begin VB.Menu mst_ro_type 
         Caption         =   "Round Type Master"
      End
   End
   Begin VB.Menu pms_trn 
      Caption         =   "&Transaction"
      Begin VB.Menu mst_rg 
         Caption         =   "Registration"
      End
      Begin VB.Menu mst_att 
         Caption         =   "&Attendance"
      End
      Begin VB.Menu mst_sel 
         Caption         =   "&Selected"
      End
   End
End
Attribute VB_Name = "pms_mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public t_count As Integer
Dim t_frm() As Form
Dim chk_open_v As Boolean
Dim alltabexit As Boolean
Public tabclose_flag As Boolean  ' for tab close




Private Sub a_Click()

End Sub

Private Sub alltabc_cmd_Click()
    Dim ans As Integer
    ans = MsgBox("Do you Want To Close All Open Tabs ?", vbYesNo + vbQuestion)
    If ans = 6 Then '6= yes
    alltabexit = True
    Dim i As Integer
    For i = 0 To t_count - 1
        Unload t_frm(i)
    Next
    
    For i = 0 To t_count - 1
      frmtab.TabVisible(i) = False
    Next
     
    ReDim t_frm(0)
    t_count = 0
    chk_open_v = False
    Picture4.Height = 0
    alltabexit = False
    End If
    
End Sub

Public Sub c_tab_cmd_Click(Index As Integer)
        Call tab_close
        If t_count = 0 Then
            Picture4.Height = 0
        End If
        
End Sub





Private Sub co_mst_mnu_fr_Click()
       Call co_mst_mnu_img_Click
End Sub






Private Sub co_mst_mnu_fr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If co_mst_mnu_fr.BorderStyle = 0 Then
        co_mst_mnu_fr.BorderStyle = 1
    End If
End Sub




Private Sub co_mst_mnu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call co_mst_mnu_fr_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If usr_win_pb.Visible = False Then
        usr_win_pb.Visible = True
    End If
    
End Sub

Private Sub frmtab_Click(PreviousTab As Integer)
   If tabclose_flag = False Then
    If alltabexit = False Then
        Call shfrm
    End If
    End If
End Sub



Private Sub co_mst_mnu_lbl_Click()
     Call co_mst_mnu_img_Click
End Sub


Private Sub Image1_Click()
        Picture1.Picture = LoadPicture(App.Path & "\loading_jpg.jpg")
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call std_mst_mnu_fr_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call pd_mst_mnu_fr_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub MDIForm_Load()
    Dim i As Integer
    
    menu.Width = Picture1.Width
    chk_open_v = False
    
     For i = 0 To frmtab.Tabs - 1
      frmtab.TabVisible(i) = False
     Next
     
    menu.TabCaption(0) = ""
    menu.TabCaption(1) = ""
    menu.TabCaption(2) = ""
    Picture4.Height = 0
    alltabexit = False
    
    'Image1.Picture = LoadPicture(App.Path & "\loading_jpg.jpg")
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (co_mst_mnu_fr.BorderStyle = 1) Then
        co_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (pd_mst_mnu_fr.BorderStyle = 1) Then
         pd_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (std_mst_mnu_fr.BorderStyle = 1) Then
         std_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (ro_typ_mst_mnu_fr.BorderStyle = 1) Then
         ro_typ_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (plc_ro_mst_mnu_fr.BorderStyle = 1) Then
      plc_ro_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (usr_win_pb.Visible = True) Then
        usr_win_pb.Visible = False
    End If
End Sub

Private Sub MDIForm_Resize()
    menu.Width = Picture1.Width

End Sub



Private Sub menu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (co_mst_mnu_fr.BorderStyle = 1) Then
        co_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (pd_mst_mnu_fr.BorderStyle = 1) Then
         pd_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (std_mst_mnu_fr.BorderStyle = 1) Then
         std_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (ro_typ_mst_mnu_fr.BorderStyle = 1) Then
         ro_typ_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (plc_ro_mst_mnu_fr.BorderStyle = 1) Then
      plc_ro_mst_mnu_fr.BorderStyle = 0
    End If
    
    If (usr_win_pb.Visible = True) Then
        usr_win_pb.Visible = False
    End If
    
End Sub


Private Sub co_mst_mnu_img_Click()
      
      If t_count > 0 Then
        Call chk_open(co_frm)
      Else
        chk_open_v = False
      End If
      
      
      If chk_open_v = False Then
        Call showfrm(co_frm)
        t_count = t_count + 1
      Else
        MsgBox ("tab already open..")
      End If
      
End Sub

Private Sub mst_att_Click()

If t_count > 0 Then
        Call chk_open(att_pd_frm)
      Else
        chk_open_v = False
      End If
      
      
      If chk_open_v = False Then
        Call showfrm(att_pd_frm)
        t_count = t_count + 1
      Else
        MsgBox ("tab already open..")
      End If
End Sub

Private Sub mst_comp_Click()
com_detail_frm.Show
End Sub

Private Sub mst_rg_Click()
   If t_count > 0 Then
        Call chk_open(rg_frm)
      Else
        chk_open_v = False
      End If
      
      
      If chk_open_v = False Then
        Call showfrm(rg_frm)
        t_count = t_count + 1
      Else
        MsgBox ("tab already open..")
      End If
End Sub

Private Sub mst_ro_Click()
        plc_ro_frm.Show
End Sub

Private Sub mst_ro_type_Click()
        ro_type_frm.Show
End Sub

Private Sub mst_sel_Click()
slec_std_frm.Show
End Sub

Private Sub pd_mst_mnu_fr_Click()
        Call pd_mst_mnu_img_Click
End Sub

Private Sub pd_mst_mnu_fr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If (pd_mst_mnu_fr.BorderStyle = 0) Then
            pd_mst_mnu_fr.BorderStyle = 1
        End If
End Sub

Private Sub pd_mst_mnu_img_Click()
    If t_count > 0 Then
        Call chk_open(pd_frm)
    Else
        chk_open_v = False
    End If
      
    If chk_open_v = False Then
        Call showfrm(pd_frm)
        t_count = t_count + 1
    Else
        MsgBox ("tab already open..")
    End If
End Sub









Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        
        If (co_mst_mnu_fr.BorderStyle = 1) Then
            co_mst_mnu_fr.BorderStyle = 0
        End If
        
        If (pd_mst_mnu_fr.BorderStyle = 1) Then
             pd_mst_mnu_fr.BorderStyle = 0
        End If
        
        If (std_mst_mnu_fr.BorderStyle = 1) Then
             std_mst_mnu_fr.BorderStyle = 0
        End If
        
        If (ro_typ_mst_mnu_fr.BorderStyle = 1) Then
         ro_typ_mst_mnu_fr.BorderStyle = 0
        End If
        
        If (plc_ro_mst_mnu_fr.BorderStyle = 1) Then
          plc_ro_mst_mnu_fr.BorderStyle = 0
        End If
        
        If (usr_win_pb.Visible = True) Then
            usr_win_pb.Visible = False
        End If
        
    
        
End Sub










Private Sub plc_ro_mst_mnu_fr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         If (plc_ro_mst_mnu_fr.BorderStyle = 0) Then
        plc_ro_mst_mnu_fr.BorderStyle = 1
        End If
End Sub

Private Sub plc_ro_mst_mnu_img_Click()
    If t_count > 0 Then
        Call chk_open(plc_ro_frm)
    Else
        chk_open_v = False
    End If
      
    If chk_open_v = False Then
        Call showfrm(plc_ro_frm)
        t_count = t_count + 1
    Else
        MsgBox ("tab already open..")
    End If
End Sub

Private Sub ro_typ_mst_mnu_fr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
         If (ro_typ_mst_mnu_fr.BorderStyle = 0) Then
         ro_typ_mst_mnu_fr.BorderStyle = 1
        End If
End Sub

Private Sub ro_typ_mst_mnu_img_Click()
    If t_count > 0 Then
        Call chk_open(ro_type_frm)
    Else
        chk_open_v = False
    End If
      
    If chk_open_v = False Then
        Call showfrm(ro_type_frm)
        t_count = t_count + 1
    Else
        MsgBox ("tab already open..")
    End If
End Sub

Private Sub std_mst_mnu_fr_Click()
        Call std_mst_mnu_img_Click
End Sub

Private Sub std_mst_mnu_fr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If (std_mst_mnu_fr.BorderStyle = 0) Then
         std_mst_mnu_fr.BorderStyle = 1
        End If
End Sub

Private Sub std_mst_mnu_img_Click()

    If t_count > 0 Then
        Call chk_open(std_frm)
    Else
        chk_open_v = False
    End If
      
    If chk_open_v = False Then
        Call showfrm(std_frm)
        t_count = t_count + 1
    Else
        MsgBox ("tab already open..")
    End If

End Sub

Private Sub chk_open(frm As Form)
    Dim i As Integer
    
    For i = 0 To t_count - 1
    If (t_frm(i).Name = frm.Name) Then
        chk_open_v = True
        Exit For
    Else
        chk_open_v = False
    End If
    Next
    
End Sub

Private Sub showfrm(frm As Form)
    ReDim Preserve t_frm(t_count)
    
    Set t_frm(t_count) = frm
    frmtab.TabVisible(t_count) = True
    
    frm.Show
    
    frm.Top = 0
    frm.Left = 0
    frmtab.Tab = t_count
    frmtab.Caption = frm.Caption
    
    If Picture4.Height = 15 Then
        Picture4.Height = 550
    End If
    
   ' MsgBox t_frm(t_count).Caption
    
    
End Sub

Private Sub shfrm()
   
        Dim i As Integer
        
        For i = 0 To t_count - 1
            If (i = frmtab.Tab) Then
                t_frm(i).Visible = True
            Else
               t_frm(i).Visible = False
            End If
        Next
End Sub


Public Sub tab_close()
    tabclose_flag = True
    Unload t_frm(frmtab.Tab)
    
    Dim i As Integer
     For i = frmtab.Tab To UBound(t_frm) - 1
        Set t_frm(i) = t_frm(i + 1)
    Next
    
    For i = frmtab.Tab To t_count
        frmtab.TabCaption(i) = frmtab.TabCaption(i + 1)
    Next
    
    frmtab.TabVisible(t_count - 1) = False

    
    t_count = t_count - 1
    
    If t_count <> 0 Then
        ReDim Preserve t_frm(UBound(t_frm) - 1)
        t_frm(frmtab.Tab).Visible = True
    Else
        ReDim t_frm(0)
    End If
    
    tabclose_flag = False
    
End Sub

