VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      Picture         =   "Form1.frx":B4AE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Picture         =   "Form1.frx":BBC2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton clr_cmd 
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Picture         =   "Form1.frx":C18F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":C54F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   6000
      Picture         =   "Form1.frx":CA7C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   6
      BackColor       =   12648447
      BackColorFixed  =   8438015
      BackColorBkg    =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

