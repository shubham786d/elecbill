VERSION 5.00
Begin VB.Form loadinggen_frm 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form3"
   ScaleHeight     =   2310
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer labelTimer 
      Interval        =   750
      Left            =   4920
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2880
      Top             =   480
   End
   Begin VB.Timer timeout_tmr 
      Interval        =   7500
      Left            =   720
      Top             =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1290
      Left            =   2400
      Picture         =   "loadinggen_frm.frx":0000
      Top             =   0
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "                     Generating Bills..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "loadinggen_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim lablecount As Integer

Private Sub labelTimer_Timer()
    Select Case lablecount
        Case 0
                Label2.Caption = "Calculating Energy Charges..."
        Case 1
                Label2.Caption = "Calculating Fixed Charges..."
        Case 2
                Label2.Caption = "Calculating Electricity charges..."
        Case 3
                Label2.Caption = "Calculating Load Tax..."
        Case 4
                Label2.Caption = "Calculating Meter Rent..."
        Case 5
                Label2.Caption = "Calculating Security Charges..."
        Case 6
                Label2.Caption = "Calculating Interest on subsity..."
        Case 7
                Label2.Caption = "Calculating Subsity..."
        Case 8
                Label2.Caption = "Calculating Total Bill..."
    End Select
    lablecount = lablecount + 1
End Sub

Private Sub timeout_tmr_Timer()
Unload Me
End Sub

Private Sub Timer1_Timer()
Debug.Print "" & App.Path & "" & "\img\gen loading imag\frame_" & i & "_delay-0.04s.gif"
Image1.Picture = LoadPicture("" & App.Path & "" & "\img\gen loading imag\frame_" & i & "_delay-0.04s.gif")
   
    i = i + 1
    If i > 19 Then
        i = 0
    End If
End Sub
