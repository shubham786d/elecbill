VERSION 5.00
Begin VB.Form cnsu_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consumer Form"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Palette         =   "cnsu_frm.frx":0000
   Picture         =   "cnsu_frm.frx":B4AE
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame form_frame 
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   8295
      Begin VB.CheckBox email_chk 
         BackColor       =   &H8000000E&
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox emailid_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         MaxLength       =   255
         TabIndex        =   20
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox cname_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         MaxLength       =   255
         TabIndex        =   19
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox mob_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox phnno_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2760
         Width           =   2895
      End
      Begin VB.CheckBox mob_chk 
         BackColor       =   &H8000000E&
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox phn_chk 
         BackColor       =   &H8000000E&
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton src_cmd 
         Caption         =   "Search"
         Height          =   735
         Left            =   7200
         Picture         =   "cnsu_frm.frx":20265
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton s_cmd 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         Picture         =   "cnsu_frm.frx":208D7
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton del_cmd 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         Picture         =   "cnsu_frm.frx":20FEB
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton new_cmd 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         Picture         =   "cnsu_frm.frx":215B8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton co_ex_cmd 
         Height          =   375
         Left            =   5400
         Picture         =   "cnsu_frm.frx":21AE5
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone  Number"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   5655
         Left            =   0
         Top             =   120
         Width           =   8295
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         Height          =   5895
         Left            =   0
         Top             =   0
         Width           =   8295
      End
   End
   Begin VB.Frame exit_frme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   8760
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
      Begin VB.Image exit_img 
         Height          =   405
         Left            =   240
         Picture         =   "cnsu_frm.frx":21FD3
         Top             =   240
         Width           =   360
      End
      Begin VB.Label exit_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "EXIT"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame del_frme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7800
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
      Begin VB.Image del_img 
         Height          =   480
         Left            =   120
         Picture         =   "cnsu_frm.frx":227AE
         Top             =   120
         Width           =   480
      End
      Begin VB.Label del_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "DELETE"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame save_frme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6840
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
      Begin VB.Image save_img 
         Height          =   480
         Left            =   240
         Picture         =   "cnsu_frm.frx":23078
         Top             =   120
         Width           =   480
      End
      Begin VB.Label save_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "SAVE"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame new_frme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5880
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
      Begin VB.Image new_img 
         Height          =   480
         Left            =   240
         Picture         =   "cnsu_frm.frx":233BB
         Top             =   120
         Width           =   480
      End
      Begin VB.Label new_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "NEW"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame search_frme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   -21360
      Picture         =   "cnsu_frm.frx":23470
      Top             =   -11040
      Width           =   360
   End
End
Attribute VB_Name = "cnsu_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd_cnsu As ADODB.Command
Dim rs_cnsu As ADODB.Recordset
Public consumer_id As Long
Public state As Integer ' 1:insert 2:update

Private Sub Command1_Click()

End Sub

Private Sub cname_txt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case 48 To 57 'numaric
                KeyAscii = 0
    End Select
End Sub

Private Sub co_ex_cmd_Click()
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub del_cmd_Click()
If state = 2 Then
        Dim test As Integer
        test = MsgBox("Do U Want To Delete This Record ?", vbYesNoCancel + vbQuestion, "Information")
         
         If test = 6 Then
            Dim qry As String
            qry = "delete from consumer_t where cid=" & cnsu_src_frm.cnsu_id & " "
            Call delete(qry)
                   
            MsgBox "Round Type Record Deleted Successfully", vbInformation
          
            state = 3
            
            cname_txt.Text = ""
            mob_txt.Text = ""
            phnno_txt.Text = ""
            add_txt.Text = ""
            emailid_txt.Text = ""
            
            cname_txt.Enabled = False
            mob_txt.Enabled = False
            phnno_txt.Enabled = False
            add_txt.Enabled = False
            emailid_txt.Enabled = False
            
            s_cmd.Enabled = False
           ' clr_cmd.Enabled = False
            del_cmd.Enabled = False
            
            mob_chk.Enabled = False
            phn_chk.Enabled = False
            
            mob_chk.value = False
            phn_chk.value = False
        End If
End If
End Sub

Private Sub del_frme_DragDrop(Source As Control, X As Single, Y As Single)
Call del_cmd_Click
End Sub

Private Sub del_img_Click()
Call del_cmd_Click
End Sub

Private Sub del_lbl_Click()
Call del_cmd_Click
End Sub

Public Sub email_chk_Click()
    emailid_txt.Text = ""

    If email_chk.value = 1 Then
        emailid_txt.Enabled = True
    Else
        emailid_txt.Enabled = False
    End If
End Sub

Private Sub exit_frme_DragDrop(Source As Control, X As Single, Y As Single)
Call co_ex_cmd_Click
End Sub

Private Sub exit_img_Click()
Call co_ex_cmd_Click
End Sub

Private Sub exit_lbl_Click()
Call co_ex_cmd_Click
End Sub

Private Sub Form_Initialize()
    
    Call form_borderset(form_frame)
End Sub

Private Sub Form_Load()
    Set rs_cnsu = New ADODB.Recordset
    rs_cnsu.CursorLocation = adUseClient
       
    rs_cnsu.Open ("select max(cid) from consumer_t"), bms_cn, 3, 3
   ' rs_cnsu
    
    If rs_cnsu.RecordCount > 0 Then
         If IsNull(rs_cnsu.Fields(0)) Then
             consumer_id = 1
         Else
             consumer_id = rs_cnsu.Fields(0) + 1
         End If
    End If
    
    rs_cnsu.Close
    Call form_borderset(form_frame)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If (bms_mdi.tabclose_flag = False) And (bms_mdi.alltabexit = False) Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then         '// tab bar height shift to 0 if last tab//
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub Frame7_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label12_Click()

End Sub

Public Sub mob_chk_Click()
If mob_chk = 1 Then
    mob_txt.Enabled = True
Else
    mob_txt.Enabled = False
End If

mob_txt.Text = ""

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
        
        cname_txt.Enabled = True
       ' add_txt.Enabled = True
        emailid_txt.Enabled = True
        mob_chk.Enabled = True
        Call mob_chk_Click
        mob_chk.value = 1
        phn_chk.Enabled = True
        state = 1
        
        cname_txt.Text = ""
        mob_txt.Text = ""
        phnno_txt.Text = ""
       ' add_txt.Text = ""
        emailid_txt.Text = ""
End Sub

Private Sub new_frme_DragDrop(Source As Control, X As Single, Y As Single)
Call new_cmd_Click
End Sub

Private Sub new_img_Click()
Call new_cmd_Click
End Sub

Private Sub new_lbl_Click()
Call new_cmd_Click
End Sub

Private Sub Option1_Click()

End Sub

Public Sub phn_chk_Click()
phnno_txt.Text = ""

If phn_chk.value = 1 Then
    phnno_txt.Enabled = True
Else
    phnno_txt.Enabled = False
End If

End Sub

Private Sub phnno_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub s_cmd_Click()
    
            Select Case state
                Case 1
                    If cname_txt.Text <> "" Then
                                If mob_chk.value = 0 And phn_chk.value = 0 Then
                                    MsgBox "Please Input Mobile Number Or Number ", vbInformation
                                    Exit Sub
                                End If
                                
                                If mob_chk.value = 1 Then
                                    If Len(mob_txt.Text) <> 0 Then
                                        If Len(mob_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Mobile number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please Input Mobile Number", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                If phn_chk.value = 1 Then
                                    If Len(phnno_txt.Text) <> 0 Then
                                        If Len(phnno_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Phone number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please input Phone number ", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                
                                If emailid_txt <> "" Then
                                     
                                     Set cmd_cnsu = New ADODB.Command
                                     cmd_cnsu.CommandType = adCmdText
                                     
                                     Dim qry As String
                                     qry = " insert into consumer_t values('" & consumer_id & "','" & cname_txt.Text & "','" & mob_txt.Text & "','" & phnno_txt.Text & "','" & emailid_txt.Text & "')"
                                     Call insert(qry)
                                    
                                     consumer_id = consumer_id + 1
                                     state = 3
                                     
                                     s_cmd.Enabled = False
                                    ' clr_cmd.Enabled = False
                                     del_cmd.Enabled = False
                                     
                                     cname_txt.Enabled = False
                                     mob_txt.Enabled = False
                                     phnno_txt.Enabled = False
                                    
                                     emailid_txt.Enabled = False
                                     
                                     mob_chk.Enabled = False
                                     phn_chk.Enabled = False
                                     email_chk.Enabled = False
                                     
                                     mob_chk.value = False
                                     phn_chk.value = False
                                     
                                     cname_txt.Text = ""
                                     mob_txt.Text = ""
                                     phnno_txt.Text = ""

                                     emailid_txt.Text = ""
                                     
                                      MsgBox "New Record Saved Successfully", vbInformation
                            Else
                                MsgBox "Please Input EMail Address", vbInformation
                            End If
                             
                    Else
                        MsgBox "Please Input Cunsumer Name", vbInformation
                    End If
                    
                Case 2
                    If cname_txt.Text <> "" Then
                                If mob_chk.value = 0 And phn_chk.value = 0 Then
                                    MsgBox "Please Input Mobile Number Or Number ", vbInformation
                                    Exit Sub
                                End If
                                
                                If mob_chk.value = 1 Then
                                    If Len(mob_txt.Text) <> 0 Then
                                        If Len(mob_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Mobile number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please Input Mobile number", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                If phn_chk.value = 1 Then
                                    If Len(phnno_txt.Text) <> 0 Then
                                        If Len(phnno_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Phone number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please Input Phone number", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                
                                If emailid_txt <> "" Then
                    
                                    Dim str As String
                                    str = "update consumer_t set cname='" & cname_txt.Text & "' , mobno='" & mob_txt.Text & "',phno='" & phnno_txt.Text & "',emailid='" & emailid_txt.Text & "' where cid =" & cnsu_src_frm.cnsu_id & ""
                                    Call update(str)
                                    
                                    state = 3
                                    
                                    cname_txt.Text = ""
                                    mob_txt.Text = ""
                                    phnno_txt.Text = ""
                                    
                                    emailid_txt.Text = ""
                                    
                                    cname_txt.Enabled = False
                                    mob_txt.Enabled = False
                                    phnno_txt.Enabled = False
                                    
                                    emailid_txt.Enabled = False
                                    
                                    s_cmd.Enabled = False
                                    del_cmd.Enabled = False
                                    
                                    mob_chk.Enabled = False
                                    phn_chk.Enabled = False
                                    email_chk.Enabled = False
                                    
                                    mob_chk.value = False
                                    phn_chk.value = False
                                    email_chk.value = False
                                    MsgBox "Record Update SuccessFully", vbInformation
                                Else
                                MsgBox "Please Input EMail Address", vbInformation
                            End If
                             
                        
                    Else
                        MsgBox "Please Input Cunsumer Name", vbInformation
                    End If
                        
                Case 3
            End Select
End Sub

Private Sub save_frme_DragDrop(Source As Control, X As Single, Y As Single)
Call s_cmd_Click
End Sub

Private Sub save_img_Click()
Call s_cmd_Click
End Sub

Private Sub save_lbl_Click()
Call s_cmd_Click
End Sub

Private Sub search_frme_DragDrop(Source As Control, X As Single, Y As Single)
Call src_cmd_Click
End Sub

Private Sub search_lbl_Click()
Call src_cmd_Click
End Sub

Private Sub src_cmd_Click()
cnsu_src_frm.Show
End Sub

Private Sub src_img_Click()
Call src_cmd_Click
End Sub
