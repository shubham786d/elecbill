VERSION 5.00
Begin VB.Form reader_frm 
   Caption         =   "reader"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "reader_frm.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   15120
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   9360
      Picture         =   "reader_frm.frx":B4AE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton clr_cmd 
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      Picture         =   "reader_frm.frx":B99C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      Picture         =   "reader_frm.frx":DBF1
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      Picture         =   "reader_frm.frx":E1BE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   3840
      Picture         =   "reader_frm.frx":E8D2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox address 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox mob_txt 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox readerid_txt 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mobile no"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reader name"
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reader id"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   6135
      Left            =   1800
      Top             =   1200
      Width           =   9135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   1560
      Top             =   960
      Width           =   9255
   End
End
Attribute VB_Name = "reader_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_reader As ADODB.Recordset
Public readerid As Long
Dim reader_id As Long
Public state As Integer

Private Sub Form_Load()
Set rs_reader = New ADODB.Recordset
    rs_reader.CursorLocation = adUseClient
       
    rs_reader.Open ("select max(rid) from reader_t"), bms_cn, 3, 3
    
    If rs_reader.RecordCount > 0 Then
         If IsNull(rs_reader.Fields(0)) Then
             readerid = 1
         Else
             readerid = rs_reader.Fields(0) + 1
         End If
    End If
       
    readerid_txt.Text = readerid
    rs_reader.Close

End Sub

Private Sub new_cmd_Click()
s_cmd.Enabled = True
readerid = 1
reader_name = ""
mob_txt.Text = ""
address.Text = ""
'yes_opt.Enabled = True
'no_opt.Enabled = True

'yes_opt.value = True

state = 1

End Sub
Private Sub s_cmd_Click()
    Select Case state
        Case 1
            Dim qry As String
    '        Dim opt As Boolean
'            If yes_opt.value = True Then
 '               opt = 1
 '           ElseIf no_opt.value = True Then
  '              opt = 0
   '         End If
            
                    qry = "insert into reader_t values(" & readerid & ",'" & Text2.Text & "', '" & mob_txt.Text & "'," & address & ")"
                    Call insert(qry)
            MsgBox "New Record Saved Successfully", vbInformation
            state = 3
            
            readerid = -999
            'readername_txt.Text = ""
            'mobile_no.Text = ""
           ' address.Text = ""
            
            'metertypsrc_cmd.Enabled = False
            'yes_opt.Enabled = False
            'no_opt.Enabled = False
            'yes_opt.value = False
            s_cmd.Enabled = False
        Case 2
        Case 3
    End Select
    
            

End Sub

