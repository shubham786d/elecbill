VERSION 5.00
Begin VB.Form reader_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "reader"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   15120
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   9840
      Picture         =   "Þ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      Picture         =   "Þ.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      Picture         =   "Þ.frx":0D86
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   3000
      Picture         =   "Þ.frx":1353
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   7440
      Picture         =   "Þ.frx":1880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox address_txt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox mob_txt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox readname_txt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox readerid_txt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   1800
      Top             =   1080
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
Public searchedrederid As Long
Public state As Integer

Private Sub co_ex_cmd_Click()
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub del_cmd_Click()
Dim test As Integer
        test = MsgBox("Do U Want To Delete This Record ?", vbYesNoCancel + vbQuestion, "Information")
         
         If test = 6 Then
            Dim qry As String
            qry = "delete from reader_t where rid=" & searchedrederid & " "
            Call delete(qry)
                   
            MsgBox "reader Record Deleted Successfully", vbInformation
            Call formreset
        End If
End Sub

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


Private Sub mob_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub new_cmd_Click()
s_cmd.Enabled = True
del_cmd.Enabled = False

readerid_txt.Text = readerid

readname_txt.Text = ""
mob_txt.Text = ""
address_txt.Text = ""

readname_txt.Enabled = True
mob_txt.Enabled = True
address_txt.Enabled = True
state = 1

End Sub
Private Sub s_cmd_Click()
    Select Case state
        Case 1
            Dim qry As String

            qry = "insert into reader_t values(" & readerid & ",'" & readname_txt.Text & "', '" & mob_txt.Text & "','" & address_txt & "')"
            Call insert(qry)
            MsgBox "New Record Saved Successfully", vbInformation
            readerid = readerid + 1
            Call formreset
        Case 2
            Dim str As String
            str = "update reader_t set RName='" & readname_txt.Text & "',Mobileno='" & mob_txt.Text & "',address='" & address_txt.Text & "' where rid =" & searchedrederid & ""
            Call update(str)
            
            Call formreset
        Case 3
    End Select
End Sub
Private Sub formreset()
            state = 3
            readerid_txt.Text = readerid
            readname_txt.Text = ""
            mob_txt.Text = ""
            address_txt.Text = ""
            readname_txt.Enabled = flase
            mob_txt.Enabled = False
            address_txt.Enabled = False
            s_cmd.Enabled = False
            del_cmd.Enabled = False
End Sub



Private Sub src_cmd_Click()
reader_src_frm.Show vbModal
End Sub
