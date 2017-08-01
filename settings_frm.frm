VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form settings_frm 
   Caption         =   "Form3"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form3"
   ScaleHeight     =   6315
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "settings_frm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "settings_frm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "settings_frm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   4800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   4800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   4800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Delay charge"
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "% Interest on Security Deposit/"
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "last day of bill payment"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
   End
End
Attribute VB_Name = "settings_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_Click()

End Sub

