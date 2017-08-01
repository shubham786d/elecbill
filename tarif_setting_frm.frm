VERSION 5.00
Begin VB.Form tarif_setting_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Tariff Setting Form"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "tarif_setting_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   11520
      Picture         =   "tarif_setting_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   9720
      Picture         =   "tarif_setting_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   5160
      Picture         =   "tarif_setting_frm.frx":15917
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      Picture         =   "tarif_setting_frm.frx":15E44
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      Picture         =   "tarif_setting_frm.frx":16411
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox text_txt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      MaxLength       =   255
      TabIndex        =   3
      Top             =   3600
      Width           =   3015
   End
   Begin VB.OptionButton phase_opt 
      Caption         =   "Phase"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9720
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.OptionButton purpose_opt 
      Caption         =   "Purpose"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton type_opt 
      Caption         =   "Type"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type Name:"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   2880
      Top             =   480
      Width           =   9855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   2880
      Top             =   360
      Width           =   9855
   End
End
Attribute VB_Name = "tarif_setting_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_tarif_set As ADODB.Recordset
Public type_id As Long
Public purpose_id As Long
Public phase_id As Long
Public update_id As Long
Public state As Integer ' 1:insert 2:update
Public oldtaxname As String


Private Sub co_ex_cmd_Click()
If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub del_cmd_Click()
    Dim str As String
    
    If type_opt.value = True Then
        str = "delete * from type_t where tid=" & update_id & ""
    ElseIf purpose_opt.value = True Then
        str = "delete * from perpose_t where id=" & update_id & ""
    ElseIf phase_opt.value = True Then
        str = "delete * from Phase_t where pid=" & update_id & ""
    End If
    
    Call delete(str)
     
    MsgBox "Record Deleted Successfully", vbInformation
          
    state = 3
    s_cmd.Enabled = False
   ' clr_cmd.Enabled = False
    del_cmd.Enabled = False
    
    type_opt.Enabled = False
    purpose_opt.Enabled = False
    phase_opt.Enabled = False
    text_txt.Enabled = False
    
    type_opt.value = True
    text_txt.Text = ""
    oldtaxname = ""
End Sub

Private Sub Form_Load()
    Set rs_tarif_set = New ADODB.Recordset                           ' TYPE ID
    rs_tarif_set.CursorLocation = adUseClient
       
    rs_tarif_set.Open ("select max(tid) from type_t"), bms_cn, 3, 3
   ' rs_tarif_set
   
    If rs_tarif_set.RecordCount > 0 Then
         If IsNull(rs_tarif_set.Fields(0)) Then
             type_id = 1
         Else
             type_id = rs_tarif_set.Fields(0) + 1
         End If
    End If
    
    rs_tarif_set.Close
    
    
    Set rs_tarif_set = New ADODB.Recordset                          ' purpose ID
    rs_tarif_set.CursorLocation = adUseClient
       
    rs_tarif_set.Open ("select max(id) from perpose_t"), bms_cn, 3, 3
   ' rs_tarif_set
    
    If rs_tarif_set.RecordCount > 0 Then
         If IsNull(rs_tarif_set.Fields(0)) Then
             purpose_id = 1
         Else
             purpose_id = rs_tarif_set.Fields(0) + 1
         End If
    End If
    
    rs_tarif_set.Close
    
    
    Set rs_tarif_set = New ADODB.Recordset                           'phase ID
    rs_tarif_set.CursorLocation = adUseClient
       
    rs_tarif_set.Open ("select max(pid) from phase_t"), bms_cn, 3, 3
   ' rs_tarif_set
    
    If rs_tarif_set.RecordCount > 0 Then
         If IsNull(rs_tarif_set.Fields(0)) Then
              phase_id = 1
         Else
              phase_id = rs_tarif_set.Fields(0) + 1
         End If
    End If
    
    rs_tarif_set.Close
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
        If (bms_mdi.tabclose_flag = False) And (bms_mdi.alltabexit = False) Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub new_cmd_Click()
s_cmd.Enabled = True

type_opt.Enabled = True
purpose_opt.Enabled = True
phase_opt.Enabled = True
text_txt.Enabled = True

type_opt.value = True
state = 1
End Sub

Private Sub phase_opt_Click()
Label1.Caption = "Phase Name:"
text_txt.Text = ""
End Sub

Private Sub purpose_opt_Click()
Label1.Caption = "Purpose Name:"
text_txt.Text = ""
End Sub

Private Sub s_cmd_Click()
        Select Case state
            Case 1
                If text_txt <> "" Then
                     If isdup() = False Then
                     
                         Dim qry As String
                         
                         If type_opt.value = True Then
                             qry = "insert into type_t values('" & type_id & "','" & text_txt.Text & "')"
                             type_id = type_id + 1
                         ElseIf purpose_opt.value = True Then
                             qry = "insert into perpose_t values('" & purpose_id & "','" & text_txt.Text & "')"
                             purpose_id = purpose_id + 1
                         ElseIf phase_opt.value = True Then
                             qry = "insert into phase_t values('" & phase_id & "','" & text_txt.Text & "')"
                             phase_id = phase_id + 1
                         End If
                         
                         Call insert(qry)
                         MsgBox "New Record Saved Successfully", vbInformation
                         s_cmd.Enabled = False
                         text_txt.Text = ""
                         oldtaxname = ""
                         type_opt.value = True
                         type_opt.Enabled = False
                         purpose_opt.Enabled = False
                         phase_opt.Enabled = False
                         text_txt.Enabled = False
                    Else
                         MsgBox "This Name Is Already Used ", vbInformation
                    End If
                Else
                         MsgBox "Please input Some Text In field  ", vbInformation
                End If
            Case 2
                If text_txt <> "" Then
                    If isdup() = False Then
                        Dim str As String
                        If type_opt.value = True Then
                        str = "update type_t set tname='" & text_txt.Text & "' where tid=" & update_id & ""
                        ElseIf purpose_opt.value = True Then
                        str = "update perpose_t set pname='" & text_txt.Text & "' where id=" & update_id & ""
                        ElseIf phase_opt.value = True Then
                        str = "update phase_t set pname='" & text_txt.Text & "' where pid=" & update_id & ""
                        End If
                        Call update(str)
                        
                        state = 3
                        s_cmd.Enabled = False
                       ' clr_cmd.Enabled = False
                        del_cmd.Enabled = False
                        
                        type_opt.Enabled = False
                        purpose_opt.Enabled = False
                        phase_opt.Enabled = False
                        text_txt.Enabled = False
                        
                        text_txt.Text = ""
                        oldtaxname = ""
                        type_opt.value = True
                        
                        MsgBox "Record Update SuccessFully", vbInformation
                    Else
                        MsgBox "This Name Is Already Used ", vbInformation
                    End If
                Else
                         MsgBox "Please input Some Text In field  ", vbInformation
                End If
            
        End Select
End Sub

Private Sub src_cmd_Click()
tarif_setting_src_frm.Show vbModal
End Sub

Private Sub type_opt_Click()
Label1.Caption = "Type Name:"
text_txt.Text = ""
End Sub

Private Function isdup() As Boolean
    
    Set rs_tarif_set = New ADODB.Recordset
    rs_tarif_set.CursorLocation = adUseClient
    
    If oldtaxname <> text_txt.Text Then
        Dim qry As String
        If type_opt.value = True Then
            qry = "select * from type_t where tname='" & text_txt.Text & "'"
        ElseIf purpose_opt.value = True Then
            qry = "select * from perpose_t where pname='" & text_txt.Text & "'"
        ElseIf phase_opt.value = True Then
            qry = "select * from phase_t where pname='" & text_txt.Text & "'"
        End If
    
        rs_tarif_set.Open (qry), bms_cn, 3, 3

        If rs_tarif_set.RecordCount > 0 Then
            isdup = True
        ElseIf rs_tarif_set.RecordCount = 0 Then
            isdup = False
        End If
    Else
        isdup = True
    End If
End Function
