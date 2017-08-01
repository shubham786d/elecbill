VERSION 5.00
Begin VB.Form loading_frm 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timeout_tmr 
      Interval        =   9000
      Left            =   1200
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   360
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   8640
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   0
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Imagelogo 
      Height          =   1725
      Left            =   3840
      Picture         =   "loading_frm.frx":0000
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   600
      Picture         =   "loading_frm.frx":1C74
      Top             =   1920
      Width           =   4425
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   4800
      Picture         =   "loading_frm.frx":30E1
      Top             =   1920
      Width           =   3990
   End
   Begin VB.Image Image4 
      Height          =   900
      Left            =   3360
      Picture         =   "loading_frm.frx":42E4
      Top             =   2640
      Width           =   2370
   End
   Begin VB.Image Image1 
      Height          =   5490
      Left            =   600
      Picture         =   "loading_frm.frx":4E8E
      Top             =   1080
      Width           =   8250
   End
End
Attribute VB_Name = "loading_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub timeout_tmr_Timer()
    Unload Me
    login3_frm.Show
End Sub

Private Sub Timer1_Timer()
    Image1.Picture = LoadPicture("" & App.Path & "" & "\img\loding imgs\frame_" & i & "_delay-0.5s.gif")
    i = i + 1
    If i > 7 Then
        i = 0
    End If
End Sub

Private Sub Timer2_Timer()

End Sub
