VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm bms_mdi 
   BackColor       =   &H0080FFFF&
   Caption         =   "Electricity Bill Management System"
   ClientHeight    =   10650
   ClientLeft      =   -5535
   ClientTop       =   -1830
   ClientWidth     =   15120
   LinkTopic       =   "MDIForm1"
   Picture         =   "bms_mdi.frx":0000
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox logo_picbox 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7920
      Left            =   0
      Picture         =   "bms_mdi.frx":814A
      ScaleHeight     =   7890
      ScaleWidth      =   15090
      TabIndex        =   17
      Top             =   2895
      Width           =   15120
      Begin VB.Timer timeroflogoeffect 
         Interval        =   500
         Left            =   3960
         Top             =   600
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   3360
         Top             =   600
      End
      Begin VB.Image Image7 
         Height          =   4275
         Left            =   -1320
         Picture         =   "bms_mdi.frx":135F8
         Top             =   2520
         Width           =   7950
      End
      Begin VB.Image Image6 
         Height          =   4275
         Left            =   8400
         Picture         =   "bms_mdi.frx":14D5E
         Top             =   2520
         Width           =   7950
      End
      Begin VB.Image Image4 
         Height          =   900
         Left            =   5520
         Picture         =   "bms_mdi.frx":164E6
         Top             =   4200
         Width           =   2370
      End
      Begin VB.Image Image3 
         Height          =   900
         Left            =   6840
         Picture         =   "bms_mdi.frx":17090
         Top             =   3000
         Width           =   3990
      End
      Begin VB.Image Image2 
         Height          =   900
         Left            =   2400
         Picture         =   "bms_mdi.frx":18293
         Top             =   3000
         Width           =   4425
      End
      Begin VB.Image Imagelogo 
         Appearance      =   0  'Flat
         Height          =   3510
         Left            =   6360
         Picture         =   "bms_mdi.frx":19700
         Top             =   0
         Width           =   2250
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   15120
      TabIndex        =   3
      Top             =   2175
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
         Top             =   0
         Width           =   255
      End
      Begin TabDlg.SSTab frmtab 
         Height          =   6255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   19320
         _ExtentX        =   34078
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   20
         Tab             =   12
         TabsPerRow      =   10
         TabHeight       =   794
         TabMaxWidth     =   3351
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "bms_mdi.frx":1A61D
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "c_tab_cmd(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "bms_mdi.frx":1A639
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "c_tab_cmd(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "bms_mdi.frx":1A655
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "c_tab_cmd(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Tab 3"
         TabPicture(3)   =   "bms_mdi.frx":1A671
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "c_tab_cmd(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Tab 4"
         TabPicture(4)   =   "bms_mdi.frx":1A68D
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "c_tab_cmd(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Tab 5"
         TabPicture(5)   =   "bms_mdi.frx":1A6A9
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "c_tab_cmd(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Tab 6"
         TabPicture(6)   =   "bms_mdi.frx":1A6C5
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "c_tab_cmd(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "Tab 7"
         TabPicture(7)   =   "bms_mdi.frx":1A6E1
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "c_tab_cmd(7)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Tab 8"
         Tab(8).ControlEnabled=   0   'False
         Tab(8).ControlCount=   0
         TabCaption(9)   =   "Tab 9"
         Tab(9).ControlEnabled=   0   'False
         Tab(9).ControlCount=   0
         TabCaption(10)  =   "Tab 10"
         TabPicture(10)  =   "bms_mdi.frx":1A6FD
         Tab(10).ControlEnabled=   0   'False
         Tab(10).ControlCount=   0
         TabCaption(11)  =   "Tab 11"
         TabPicture(11)  =   "bms_mdi.frx":1A719
         Tab(11).ControlEnabled=   0   'False
         Tab(11).ControlCount=   0
         TabCaption(12)  =   "Tab 12"
         TabPicture(12)  =   "bms_mdi.frx":1A735
         Tab(12).ControlEnabled=   -1  'True
         Tab(12).ControlCount=   0
         TabCaption(13)  =   "Tab 13"
         TabPicture(13)  =   "bms_mdi.frx":1A751
         Tab(13).ControlEnabled=   0   'False
         Tab(13).ControlCount=   0
         TabCaption(14)  =   "Tab 14"
         TabPicture(14)  =   "bms_mdi.frx":1A76D
         Tab(14).ControlEnabled=   0   'False
         Tab(14).ControlCount=   0
         TabCaption(15)  =   "Tab 15"
         TabPicture(15)  =   "bms_mdi.frx":1A789
         Tab(15).ControlEnabled=   0   'False
         Tab(15).ControlCount=   0
         TabCaption(16)  =   "Tab 16"
         Tab(16).ControlEnabled=   0   'False
         Tab(16).ControlCount=   0
         TabCaption(17)  =   "Tab 17"
         Tab(17).ControlEnabled=   0   'False
         Tab(17).ControlCount=   0
         TabCaption(18)  =   "Tab 18"
         Tab(18).ControlEnabled=   0   'False
         Tab(18).ControlCount=   0
         TabCaption(19)  =   "Tab 19"
         Tab(19).ControlEnabled=   0   'False
         Tab(19).ControlCount=   0
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   7
            Left            =   -60360
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   6
            Left            =   -61920
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   5
            Left            =   -63720
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton c_tab_cmd 
            BackColor       =   &H00FF8080&
            Caption         =   "X"
            Height          =   195
            Index           =   4
            Left            =   -65640
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   120
            Width           =   255
         End
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
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   15060
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6120
         ScaleHeight     =   375
         ScaleWidth      =   6375
         TabIndex        =   35
         Top             =   0
         Width           =   6375
         Begin VB.TextBox welcomestrip_txt 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Baskerville Old Face"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   8175
         End
      End
      Begin VB.CommandButton user_name_cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   13080
         MaskColor       =   &H0080C0FF&
         Picture         =   "bms_mdi.frx":1A7A5
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   2055
      End
      Begin VB.PictureBox usr_win_pbb 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   12600
         ScaleHeight     =   1455
         ScaleWidth      =   2535
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "Log Out"
            Height          =   375
            Left            =   480
            TabIndex        =   37
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Exit"
            Height          =   375
            Left            =   480
            TabIndex        =   14
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "About"
            Height          =   375
            Left            =   480
            TabIndex        =   13
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Settings"
            Height          =   375
            Left            =   480
            TabIndex        =   12
            Top             =   0
            Width           =   2175
         End
         Begin VB.Image Image10 
            Height          =   375
            Left            =   25
            Picture         =   "bms_mdi.frx":2F55C
            Top             =   360
            Width           =   375
         End
         Begin VB.Image Image9 
            Height          =   375
            Left            =   25
            Picture         =   "bms_mdi.frx":2F710
            Top             =   720
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   375
            Left            =   25
            Picture         =   "bms_mdi.frx":2F78E
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   25
            Picture         =   "bms_mdi.frx":2F819
            Top             =   0
            Width           =   375
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
      Begin TabDlg.SSTab menu 
         Height          =   2175
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   54495
         _ExtentX        =   96123
         _ExtentY        =   3836
         _Version        =   393216
         Tab             =   1
         TabHeight       =   706
         TabMaxWidth     =   5292
         WordWrap        =   0   'False
         BackColor       =   16777215
         ForeColor       =   16777215
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "bms_mdi.frx":2F89B
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Shape1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Shape8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "master_mnu_back_pic(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "tarf_mnu_cmd"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "metertyp_mnu_cmd"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "meter_mnu_cmd"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "tarifset_mnu_cmd"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Picture5"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "bms_mdi.frx":3045B
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Shape2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Shape5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Shape7"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Picture9"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "read_mnu_cmd"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Con_mnu_cmd"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "showbill_mnu_cmd"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "usercreate_mnu_cmd"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Picture3"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "bms_mdi.frx":31134
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Image8"
         Tab(2).Control(1)=   "Shape6"
         Tab(2).Control(2)=   "Shape4"
         Tab(2).ControlCount=   3
         Begin VB.PictureBox Picture5 
            Height          =   495
            Left            =   -62400
            Picture         =   "bms_mdi.frx":31C8E
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   34
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture3 
            Height          =   495
            Left            =   12600
            Picture         =   "bms_mdi.frx":320BE
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   33
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton usercreate_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "User Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   7080
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":324EE
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton showbill_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Bill Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   5520
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":333B8
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Con_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Connection Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   840
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":34407
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton read_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reading Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   2400
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":350CD
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton tarifset_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "                                  Tariff Setting Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -70920
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":35D19
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton meter_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Meter Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -67800
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":36A65
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton metertyp_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Meter Type Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -66120
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":37B32
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton tarf_mnu_cmd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tariff Form"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -72480
            MaskColor       =   &H00FFFF00&
            Picture         =   "bms_mdi.frx":38B98
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Width           =   1335
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   0
            Picture         =   "bms_mdi.frx":39774
            ScaleHeight     =   1785
            ScaleWidth      =   15465
            TabIndex        =   16
            Top             =   480
            Width           =   15495
            Begin VB.CommandButton billgen_mnu_cmd 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Bill Generation Form"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   3960
               MaskColor       =   &H00FFFF00&
               Picture         =   "bms_mdi.frx":3B9C9
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.PictureBox master_mnu_back_pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Index           =   0
            Left            =   -75000
            Picture         =   "bms_mdi.frx":3C3A0
            ScaleHeight     =   1785
            ScaleWidth      =   15225
            TabIndex        =   15
            Top             =   480
            Width           =   15255
            Begin VB.CommandButton tax_mnu_cmd 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Tax Type Form"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   5640
               MaskColor       =   &H00FFFF00&
               Picture         =   "bms_mdi.frx":3E5F5
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   120
               Width           =   1335
            End
            Begin VB.CommandButton cun_mnu_cmd 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Consumer Form"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1335
               Left            =   960
               MaskColor       =   &H00FFFF00&
               Picture         =   "bms_mdi.frx":3EF4A
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Image Image8 
            Height          =   1800
            Left            =   -75000
            Picture         =   "bms_mdi.frx":3FC80
            Top             =   480
            Width           =   18930
         End
         Begin VB.Shape Shape8 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   45
            Left            =   -75000
            Top             =   430
            Width           =   3015
         End
         Begin VB.Shape Shape7 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   45
            Left            =   3120
            Top             =   435
            Width           =   3015
         End
         Begin VB.Shape Shape6 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000D&
            FillColor       =   &H80000005&
            Height          =   45
            Left            =   -68880
            Top             =   430
            Width           =   3015
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
   Begin VB.Menu tarifmnu 
      Caption         =   "&Tarif"
   End
   Begin VB.Menu cnsumnu 
      Caption         =   "&Cunsumer"
   End
   Begin VB.Menu taxtypemnu 
      Caption         =   "&Tax Type"
   End
   Begin VB.Menu metertypmnu 
      Caption         =   "&Meter Type"
   End
   Begin VB.Menu tarifsetmnu 
      Caption         =   "&Tarif Setting"
   End
   Begin VB.Menu tariftaxmnu 
      Caption         =   "&Tarif Tax "
   End
   Begin VB.Menu metermnu 
      Caption         =   "&Meter"
   End
   Begin VB.Menu conmnu 
      Caption         =   "&Connection"
   End
   Begin VB.Menu readermnu 
      Caption         =   "reader_frm"
   End
   Begin VB.Menu reading_frmmnu 
      Caption         =   "reading"
   End
   Begin VB.Menu bill_frm_mnu 
      Caption         =   "BIll "
   End
   Begin VB.Menu showbilll 
      Caption         =   "show BILL"
   End
   Begin VB.Menu paybillmenu 
      Caption         =   "pay bill"
   End
   Begin VB.Menu billgenn 
      Caption         =   "billgenn"
   End
End
Attribute VB_Name = "bms_mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As ADODB.Recordset
Public t_count As Integer
Dim t_frm() As Form
Dim chk_open_v As Boolean
Public alltabexit As Boolean
Public tabclose_flag As Boolean  ' for tab close
Dim menu_icon_color As Long
Dim menu_icon_click_color As Long
Dim logoi As Integer
Public loginid As String 'log in id
Public username As String
Dim userright As Integer

Private Sub icon_back_reset() ' will reset back of icons
    '//master forms
    If cun_mnu_cmd.BackColor = menu_icon_click_color Then
        cun_mnu_cmd.BackColor = menu_icon_color
        cun_mnu_cmd.Move cun_mnu_cmd.Left, cun_mnu_cmd.Top - 100, cun_mnu_cmd.Width, cun_mnu_cmd.Width
    End If

    If tarf_mnu_cmd.BackColor = menu_icon_click_color Then
        tarf_mnu_cmd.BackColor = menu_icon_color
        tarf_mnu_cmd.Move tarf_mnu_cmd.Left, tarf_mnu_cmd.Top - 100, tarf_mnu_cmd.Width, tarf_mnu_cmd.Width
    End If

    If tarifset_mnu_cmd.BackColor = menu_icon_click_color Then
        tarifset_mnu_cmd.BackColor = menu_icon_color
        
        tarifset_mnu_cmd.Move tarifset_mnu_cmd.Left, tarifset_mnu_cmd.Top - 100, tarifset_mnu_cmd.Width, tarifset_mnu_cmd.Width
    End If

    If tax_mnu_cmd.BackColor = menu_icon_click_color Then
        tax_mnu_cmd.BackColor = menu_icon_color
        tax_mnu_cmd.Move tax_mnu_cmd.Left, tax_mnu_cmd.Top - 100, tax_mnu_cmd.Width, tax_mnu_cmd.Width
    End If

    If metertyp_mnu_cmd.BackColor = menu_icon_click_color Then
        metertyp_mnu_cmd.BackColor = menu_icon_color
        metertyp_mnu_cmd.Move metertyp_mnu_cmd.Left, metertyp_mnu_cmd.Top - 100, metertyp_mnu_cmd.Width, metertyp_mnu_cmd.Width
    End If

    If meter_mnu_cmd.BackColor = menu_icon_click_color Then
        meter_mnu_cmd.BackColor = menu_icon_color
        meter_mnu_cmd.Move meter_mnu_cmd.Left, meter_mnu_cmd.Top - 100, meter_mnu_cmd.Width, meter_mnu_cmd.Width
    End If
    
    '// side menu
    If usr_win_pbb.Visible = True Then
         usr_win_pbb.Visible = False
    End If
    
    '// transaction form
    
    If Con_mnu_cmd.BackColor = menu_icon_click_color Then
        Con_mnu_cmd.BackColor = menu_icon_color
        Con_mnu_cmd.Move Con_mnu_cmd.Left, Con_mnu_cmd.Top - 100, Con_mnu_cmd.Width, Con_mnu_cmd.Width
    End If
    
    If read_mnu_cmd.BackColor = menu_icon_click_color Then
        read_mnu_cmd.BackColor = menu_icon_color
        read_mnu_cmd.Move read_mnu_cmd.Left, read_mnu_cmd.Top - 100, read_mnu_cmd.Width, read_mnu_cmd.Width
    End If
    
    If billgen_mnu_cmd.BackColor = menu_icon_click_color Then
        billgen_mnu_cmd.BackColor = menu_icon_color
        billgen_mnu_cmd.Move billgen_mnu_cmd.Left, billgen_mnu_cmd.Top - 100, billgen_mnu_cmd.Width, billgen_mnu_cmd.Width
    End If
    
    If showbill_mnu_cmd.BackColor = menu_icon_click_color Then
        showbill_mnu_cmd.BackColor = menu_icon_color
        showbill_mnu_cmd.Move showbill_mnu_cmd.Left, showbill_mnu_cmd.Top - 100, showbill_mnu_cmd.Width, showbill_mnu_cmd.Width
    End If
    
    If usercreate_mnu_cmd.BackColor = menu_icon_click_color Then
        usercreate_mnu_cmd.BackColor = menu_icon_color
        usercreate_mnu_cmd.Move usercreate_mnu_cmd.Left, usercreate_mnu_cmd.Top - 100, usercreate_mnu_cmd.Width, usercreate_mnu_cmd.Width
    End If
End Sub

Private Sub alltabc_cmd_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you Want To Close All Open Tabs ?", vbYesNo + vbQuestion)
    
    If ans = 6 Then '6= yes
        alltabexit = True
        Dim k As Integer
        Static Tabcount As Integer
        
        Tabcount = t_count - 1
        'MsgBox UBound(t_frm)
        For k = 0 To Tabcount
                'MsgBox UBound(t_frm)
                Unload t_frm(k)
        Next
        
        Dim i As Integer
        For i = 0 To t_count - 1
          frmtab.TabVisible(i) = False
        Next
         
        ReDim t_frm(0)
        t_count = 0
        chk_open_v = False
        Picture4.Height = 0
        Picture4.Height = 0
        logo_picbox.Visible = True
        alltabexit = False
    End If
    
End Sub

Public Sub tab_close()
    
    tabclose_flag = True
    Unload t_frm(frmtab.Tab)
    
    Dim i As Integer
     For i = frmtab.Tab To UBound(t_frm) - 1       '// tab form shift //
        Set t_frm(i) = t_frm(i + 1)
    Next
    
    For i = frmtab.Tab To t_count                  '// tab caption shift //
        frmtab.TabCaption(i) = frmtab.TabCaption(i + 1)
    Next
    
    frmtab.TabVisible(t_count - 1) = False          '//tab unshow

    
    t_count = t_count - 1
    
    If t_count <> 0 Then
        'ReDim Preserve t_frm(UBound(t_frm) - 1)
        t_frm(frmtab.Tab).Visible = True
    Else
        ReDim t_frm(0)
    End If
    
    tabclose_flag = False
End Sub


Private Sub bill_frm_mnu_Click()
Call opentab(bill_frm)
End Sub

Private Sub billgen_mnu_cmd_Click()
    Call opentab(billgen_frm)
End Sub

Private Sub billgen_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If billgen_mnu_cmd.BackColor = menu_icon_color Then
        billgen_mnu_cmd.BackColor = menu_icon_click_color
        billgen_mnu_cmd.Move billgen_mnu_cmd.Left, billgen_mnu_cmd.Top + 100, billgen_mnu_cmd.Width, billgen_mnu_cmd.Width
    End If
End Sub

Private Sub billgenn_Click()
    Call opentab(billgen_frm)
End Sub

Public Sub c_tab_cmd_Click(Index As Integer)
        Call tab_close
        If t_count = 0 Then
            Picture4.Height = 0
            logo_picbox.Visible = True
        End If
End Sub


Private Sub cnsumnu_Click()
    Call opentab(cnsu_frm)
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

Private Sub Command10_Click()

End Sub

Private Sub Command1_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you really want to log out ?", vbYesNo + vbQuestion)
    
    If ans = 6 Then '6= yes
        Unload Me
        login3_frm.Show
    End If
End Sub

Private Sub Command5_Click()
    About_frm.Show vbModal
End Sub

Private Sub Con_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Con_mnu_cmd.BackColor = menu_icon_color Then
        Con_mnu_cmd.BackColor = menu_icon_click_color
        Con_mnu_cmd.Move Con_mnu_cmd.Left, Con_mnu_cmd.Top + 100, Con_mnu_cmd.Width, Con_mnu_cmd.Width
    End If
End Sub



Private Sub master_mnu_back_pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call icon_back_reset
End Sub



Private Sub menu_Click(PreviousTab As Integer)
   
    Select Case userright
        Case 2
            If PreviousTab = 5 Then
                menu.Tab = 1
                Exit Sub
            End If
            
            
            If menu.Tab = 0 Then
                MsgBox "You Dont Have User Rights For Accesing This Tab", vbInformation
                menu.Tab = 1
            End If
        Case 1
            If PreviousTab = 5 Then
                menu.Tab = 0
                Exit Sub
            End If
            
            If menu.Tab = 1 Then
                MsgBox "You Dont Have User Rights For Accesing This Tab", vbInformation
                menu.Tab = 0
            End If
    End Select
    
    'MsgBox PreviousTab
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call icon_back_reset
End Sub


Private Sub usercreate_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If usercreate_mnu_cmd.BackColor = menu_icon_color Then
        usercreate_mnu_cmd.BackColor = menu_icon_click_color
        usercreate_mnu_cmd.Move usercreate_mnu_cmd.Left, usercreate_mnu_cmd.Top + 100, usercreate_mnu_cmd.Width, usercreate_mnu_cmd.Width
    End If
End Sub

Private Sub read_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If read_mnu_cmd.BackColor = menu_icon_color Then
        read_mnu_cmd.BackColor = menu_icon_click_color
        read_mnu_cmd.Move read_mnu_cmd.Left, read_mnu_cmd.Top + 100, read_mnu_cmd.Width, read_mnu_cmd.Width
    End If
End Sub

Private Sub showbill_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If showbill_mnu_cmd.BackColor = menu_icon_color Then
        showbill_mnu_cmd.BackColor = menu_icon_click_color
        showbill_mnu_cmd.Move showbill_mnu_cmd.Left, showbill_mnu_cmd.Top + 100, showbill_mnu_cmd.Width, showbill_mnu_cmd.Width
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
    welcomestrip_txt.Text = Right(welcomestrip_txt.Text, Len(welcomestrip_txt.Text) - 1) & Left(welcomestrip_txt.Text, 1)
End Sub

Private Sub timeroflogoeffect_Timer()
    
    logoi = logoi + 1
    Imagelogo.Picture = LoadPicture(App.Path & "\img\logo\new" & logoi & ".gif")
    If logoi = 5 Then
        logoi = 0
    End If
End Sub

Private Sub user_name_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If usr_win_pbb.Visible = False Then
        usr_win_pbb.Visible = True
    End If
    
End Sub


Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command13.BackColor = &HC0E0FF
End Sub



Private Sub Command6_Click()
    Dim ans As Integer
    
    ans = MsgBox("Do you really want to Exit ?", vbYesNo + vbQuestion)
    
    If ans = 6 Then '6= yes
        Unload Me
    End If
End Sub

Private Sub Con_mnu_cmd_Click()
    Call opentab(con_frm)
End Sub

Private Sub conmnu_Click()
    Call opentab(con_frm)
End Sub



Private Sub cun_mnu_cmd_Click()
    Call opentab(cnsu_frm)
End Sub

Private Sub cun_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cun_mnu_cmd.BackColor = menu_icon_color Then
        cun_mnu_cmd.BackColor = menu_icon_click_color
        cun_mnu_cmd.Move cun_mnu_cmd.Left, cun_mnu_cmd.Top + 100, cun_mnu_cmd.Width, cun_mnu_cmd.Width
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


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call pd_mst_mnu_fr_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub logo_picbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call icon_back_reset
End Sub


Private Sub MDIForm_Initialize()
 
     
 Image2.Move logo_picbox.Left + logo_picbox.Width / 2 - Image2.Width / 2 - 2000, logo_picbox.Top + 550 'Elc bill logo image
  
 Image3.Move Image2.Left + Image2.Width + 20, Image2.Top   'system logo image
 
 Image4.Move logo_picbox.Left + logo_picbox.Width / 2 - Image4.Width / 2, Image2.Top + Image2.Height + 250    ' system logo image
    
  
End Sub

Private Sub MDIForm_Load()
    Call checkuserright                 'check the rights of user login
    If userright <> 4 Then
        usercreate_mnu_cmd.Visible = False
    End If
    menu_Click (5)
    
    menu_icon_color = vbWhite           'color of icon fade
    menu_icon_click_color = &HC0C000
    
    
    menu.TabVisible(2) = False
    
    Dim i As Integer
    
    menu.Width = Picture1.Width
    chk_open_v = False
    
     For i = 0 To frmtab.Tabs - 1
      frmtab.TabVisible(i) = False
     Next
     
    menu.TabCaption(0) = ""
    menu.TabCaption(1) = ""
   ' menu.TabCaption(2) = ""
    Picture4.Height = 0
    alltabexit = False
    'Picture8.Width = Screen.Width
    logo_picbox.Height = Screen.Width
    
   
    Dim str As String
   ' MsgBox Hour(Time())
    If Hour(Time()) >= 12 And Hour(Time()) < 17 Then
        str = "Good After Noon "
    ElseIf Hour(Time()) > 17 And Hour(Time()) < 21 Then
        str = "Good Evening "
    ElseIf Hour(Time()) >= 5 And Hour(Time()) < 12 Then
        str = "Good Morning "
    Else
        str = "Good Night "
    End If
    
    welcomestrip_txt.Text = " Welcome And " & str & username & "....                                "
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call icon_back_reset
End Sub

Private Sub MDIForm_Resize()
    'menu.Width = Picture1.Width
    'logo_img.Left = logo_img.Width / 2 + logo_img.Width
    'logo_img.Top = logo_img.Height / 2 + logo_img.Height
End Sub



Private Sub menu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call icon_back_reset
End Sub

Private Sub meter_mnu_cmd_Click()
    Call opentab(meter_frm)
End Sub

Private Sub meter_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If meter_mnu_cmd.BackColor = menu_icon_color Then
        meter_mnu_cmd.BackColor = menu_icon_click_color
        meter_mnu_cmd.Move meter_mnu_cmd.Left, meter_mnu_cmd.Top + 100, meter_mnu_cmd.Width, meter_mnu_cmd.Width
    End If
End Sub

Private Sub metermnu_Click()
    Call opentab(meter_frm)
End Sub

Private Sub metertyp_mnu_cmd_Click()
    Call opentab(metertype_frm)
End Sub

Private Sub metertyp_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If metertyp_mnu_cmd.BackColor = menu_icon_color Then
        metertyp_mnu_cmd.BackColor = menu_icon_click_color
        metertyp_mnu_cmd.Move metertyp_mnu_cmd.Left, metertyp_mnu_cmd.Top + 100, metertyp_mnu_cmd.Width, metertyp_mnu_cmd.Width
    End If
End Sub

Private Sub metertypmnu_Click()
'metertype_frm

Call opentab(metertype_frm)

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



Private Sub paybillmenu_Click()
Call opentab(paybill_frm)
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call icon_back_reset
End Sub



Private Sub chk_open(frm As Form)
    Dim i As Integer
    
    For i = 0 To t_count - 1
    If (t_frm(i).Name = frm.Name) Then
        chk_open_v = True
        Exit For
    Else
        chk_open_v = False
        logo_picbox.Visible = False   ' logo disable
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
        Picture4.Height = 500
    End If
    
   If t_count = 0 Then
   logo_picbox.Visible = False
   End If
    
End Sub

Private Sub shfrm()  '//for showing the switing of forms in tab
   
        Dim i As Integer
        
        For i = 0 To t_count - 1
            If (i = frmtab.Tab) Then
                t_frm(i).Visible = True
            Else
               t_frm(i).Visible = False
            End If
        Next
End Sub





Private Sub read_mnu_cmd_Click()
    Call opentab(reading_frm)
End Sub

Private Sub readermnu_Click()
Call opentab(reader_frm)
End Sub

Private Sub reading_frmmnu_Click()
Call opentab(reading_frm)
End Sub

Private Sub showbill_mnu_cmd_Click()
    Call opentab(showbill_frm)
End Sub

Private Sub showbilll_Click()
Call opentab(showbill_frm)
End Sub

Private Sub tarf_mnu_cmd_Click()
Call opentab(tarif_frm)
End Sub

Private Sub tarf_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tarf_mnu_cmd.BackColor = menu_icon_color Then
        tarf_mnu_cmd.BackColor = menu_icon_click_color
        tarf_mnu_cmd.Move tarf_mnu_cmd.Left, tarf_mnu_cmd.Top + 100, tarf_mnu_cmd.Width, tarf_mnu_cmd.Width
    End If
End Sub

Private Sub tarifmnu_Click()
Call opentab(tarif_frm)
End Sub

Private Sub tarifset_mnu_cmd_Click()
Call opentab(tarif_setting_frm)
End Sub

Private Sub tarifset_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tarifset_mnu_cmd.BackColor = menu_icon_color Then
tarifset_mnu_cmd.BackColor = menu_icon_click_color
tarifset_mnu_cmd.Move tarifset_mnu_cmd.Left, tarifset_mnu_cmd.Top + 100, tarifset_mnu_cmd.Width, tarifset_mnu_cmd.Width
End If
End Sub

Private Sub tarifsetmnu_Click()
Call opentab(tarif_setting_frm)
End Sub

Private Sub tariftaxmnu_Click()
Call opentab(tariftax_frm)
End Sub

Private Sub tax_mnu_cmd_Click()
Call opentab(tax_typ_frm)
End Sub

Private Sub tax_mnu_cmd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If tax_mnu_cmd.BackColor = menu_icon_color Then
tax_mnu_cmd.BackColor = menu_icon_click_color
tax_mnu_cmd.Move tax_mnu_cmd.Left, tax_mnu_cmd.Top + 100, tax_mnu_cmd.Width, tax_mnu_cmd.Width
End If
End Sub

Private Sub taxtypemnu_Click()
Call opentab(tax_typ_frm)
End Sub


Private Sub opentab(frm As Form)
    If t_count < 8 Then
          If t_count > 0 Then
            Call chk_open(frm)
          Else
            chk_open_v = False
          End If
          
          
          If chk_open_v = False Then
            Call showfrm(frm)
            t_count = t_count + 1
          Else
            MsgBox ("tab already open..")
          End If
          Call icon_back_reset
    Else
        MsgBox "max nuber of tab is open "
    End If
    
End Sub

Private Sub usercreate_mnu_cmd_Click()
 Call opentab(createuser_frm)
End Sub

Private Sub checkuserright()
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
       
    rst.Open ("select * from login_t where userid='" & loginid & "'"), bms_cn, 3, 3
    
    If rst.RecordCount <> 0 Then
        userright = rst.Fields(3)
    Else
        userright = 4
    End If
End Sub
