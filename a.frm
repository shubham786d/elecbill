VERSION 5.00
Begin VB.Form meter_frm 
   Caption         =   "Form2"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form2"
   ScaleHeight     =   6915
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox meterid_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton clr_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      Picture         =   "a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   6840
      Picture         =   "a.frx":07B2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   840
      Picture         =   "a.frx":0E6E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      Picture         =   "a.frx":156D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Picture         =   "a.frx":1E39
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton src_cmd 
      Height          =   375
      Left            =   6000
      Picture         =   "a.frx":2604
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   2535
   End
   Begin VB.OptionButton no_opt 
      Caption         =   "No"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton yes_opt 
      Caption         =   "Yes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox meterrent_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox metertype_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Meter ID"
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Working"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Meter Rent :"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Meter Type :"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "meter_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label3_Click()

End Sub

Private Sub new_cmd_Click()

metertype_txt.Enabled = True
yes_opt.Enabled = True
no_opt.Enabled = True

End Sub
