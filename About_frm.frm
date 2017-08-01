VERSION 5.00
Begin VB.Form About_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About "
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8655
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By: Shubham Dwivedi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   4560
      Picture         =   "About_frm.frx":0000
      Top             =   240
      Width           =   3990
   End
   Begin VB.Image Image4 
      Height          =   900
      Left            =   3000
      Picture         =   "About_frm.frx":1203
      Top             =   960
      Width           =   2370
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   120
      Picture         =   "About_frm.frx":1DAD
      Top             =   240
      Width           =   4425
   End
End
Attribute VB_Name = "About_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Imagelogo_Click()

End Sub

