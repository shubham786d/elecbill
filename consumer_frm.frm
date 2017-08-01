VERSION 5.00
Begin VB.Form cnsu_frm 
   Caption         =   "Form2"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   10410
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      Picture         =   "consumer_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Picture         =   "consumer_frm.frx":07CB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      Picture         =   "consumer_frm.frx":1097
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   0
      Picture         =   "consumer_frm.frx":1849
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   6000
      Picture         =   "consumer_frm.frx":1F48
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton s_cmd 
      Height          =   375
      Left            =   1560
      Picture         =   "consumer_frm.frx":2604
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Height          =   375
      Left            =   3000
      Picture         =   "consumer_frm.frx":2DCF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton clr_cmd 
      Height          =   375
      Left            =   4560
      Picture         =   "consumer_frm.frx":369B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   0
      Picture         =   "consumer_frm.frx":3E4D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   6000
      Picture         =   "consumer_frm.frx":454C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton src_cmd 
      Height          =   375
      Left            =   7440
      Picture         =   "consumer_frm.frx":4C08
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "cnsu_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd_con As ADODB.Command
Dim rs_con As ADODB.Recordset
Public consumer_id As Long
Public state As Integer ' 1:insert 2:update

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
        cname_txt.Enabled = True
        mob_txt.Enabled = True
        phnno_txt.Enabled = True
        add_txt.Enabled = True
        state = 1
End Sub

Private Sub Command5_Click()
                    
                    Select Case state
                    Case 1
                        cmd_con.CommandType = adCmdText
                        
                        cmd_con.CommandText = " insert into consumer_t values('" & consumer_id & "','" & cname_txt.Text & "','" & mob_txt.Text & "','" & phnno_txt.Text & "','" & add_txt.Text & "')"
                        Debug.Print cmd_con.CommandText
                        cmd_con.ActiveConnection = bms_cn
                        
                        cmd_con.Execute
                    End Select
End Sub

Private Sub Form_Load()

End Sub
