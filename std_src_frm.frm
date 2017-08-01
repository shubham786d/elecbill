VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form std_src_frm 
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "std_src_frm.frx":0000
   ScaleHeight     =   6375
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid std_search_dg 
      Height          =   2655
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton clr_cmd 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox searchbox_txt 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000018&
      Caption         =   "Student Name"
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000018&
      Caption         =   "Student Scholar Number"
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton exit_cmd 
      Caption         =   "Exit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label searchLabel_lbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH STUDENT "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   5895
      Left            =   480
      Top             =   120
      Width           =   9135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   5895
      Left            =   720
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "std_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public cldform As Integer  ' calling form value
Dim rs_std_src As ADODB.Recordset
'Dim cmd_std_src As ADODB.Command
Dim flag As Boolean
Public stdschno As Long
Dim rs As ADODB.Recordset

Private Sub clr_cmd_Click()
    searchbox_txt.Text = ""
    searchbox_txt.SetFocus
End Sub

Private Sub exit_cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    flag = False
    searchbox_txt.Visible = False
    clr_cmd.Visible = False
End Sub

Private Sub Option1_Click()
    searchbox_txt.Visible = True
    clr_cmd.Visible = True
    searchLabel_lbl.Caption = "Scholar Number:"
    flag = False
    searchbox_txt.Text = ""
    searchbox_txt.MaxLength = 8
    searchbox_txt.SetFocus
End Sub

Private Sub Option2_Click()
    searchbox_txt.Visible = True
    clr_cmd.Visible = True
    searchLabel_lbl.Caption = "Student Name:"
    flag = True
    searchbox_txt.Text = ""
    searchbox_txt.MaxLength = 255
    searchbox_txt.SetFocus
End Sub

Private Sub searchbox_txt_Change()
        If searchbox_txt.Text <> "" Then
              
                Select Case cldform
                    Case 1
                            Set rs_std_src = New ADODB.Recordset
                            rs_std_src.CursorLocation = adUseClient
                            
                            If flag = True Then
                                rs_std_src.Open "select *,course_t.cname from std_t ,course_t where std_t.course = course_t.cid and s_name like '%" & searchbox_txt.Text & "%' ", pms_cn
                            ElseIf flag = False Then
                                rs_std_src.Open "select * from std_t where s_no like '" & searchbox_txt.Text & "%' ", pms_cn
                            End If
                            
                            
                    Case 2
                            Set rs_std_src = New ADODB.Recordset
                            rs_std_src.CursorLocation = adUseClient
                            
                            If flag = True Then
                                rs_std_src.Open "select * from std_t where s_name like '%" & searchbox_txt.Text & "%' and course=" & rg_frm.co_cmb.ItemData(rg_frm.co_cmb.ListIndex) & " ", pms_cn
                            ElseIf flag = False Then
                                rs_std_src.Open "select * from std_t where s_no like '" & searchbox_txt.Text & "%' and Course=" & rg_frm.co_cmb.ItemData(rg_frm.co_cmb.ListIndex) & " ", pms_cn
                            End If
                    Case 3
                             Set rs_std_src = New ADODB.Recordset
                              rs_std_src.CursorLocation = adUseClient
                              
                              If att_pd_frm.ro_cmb.ListIndex = 1 Then
                                     If flag = True Then
                                       Debug.Print "select std_t.* from std_t,place_rg_t  where std_t.s_name like '%" & searchbox_txt.Text & "%' and std_t.s_no=place_rg_t.s_no and place_rg_t.pid=" & att_pd_frm.pd_id & " "
                                        rs_std_src.Open "select std_t.* from std_t,place_rg_t  where std_t.s_name like '%" & searchbox_txt.Text & "%' and std_t.s_no=place_rg_t.s_no and place_rg_t.pid=" & att_pd_frm.pd_id & " ", pms_cn
                                    
                                    ElseIf flag = False Then
                                        Debug.Print "select std_t.* from std_t , place_rg_t where std_t .s_no like '" & searchbox_txt.Text & "%' and  std_t.s_no=place_rg_t.s_no ",
                                        rs_std_src.Open "select std_t.* from std_t , place_rg_t where std_t.s_no like '" & searchbox_txt.Text & "%' and  std_t.s_no=place_rg_t.s_no ", pms_cn
                                    End If
                             ElseIf att_pd_frm.ro_cmb.ListIndex > 1 Then
                                    Dim roundid As Long
                                    roundid = att_pd_frm.ro_cmb.ItemData(att_pd_frm.ro_cmb.ListIndex - 1)
                                    If flag = True Then
                                       Debug.Print "select std_t.* from std_t,place_rg_t  where std_t.s_name like '%" & searchbox_txt.Text & "%' and std_t.s_no=place_rg_t.s_no and place_rg_t.pid=" & att_pd_frm.pd_id & " "
                                        rs_std_src.Open "select std_t.* from std_t,std_round_t where std_t.s_name like '%" & searchbox_txt.Text & "%' and std_t.s_no=std_round_t.s_no and std_round_t.pid=" & att_pd_frm.pd_id & " and rid=" & roundid & "", pms_cn
                                    
                                    ElseIf flag = False Then
                                        Debug.Print "select std_t.* from std_t , place_rg_t where std_t .s_no like '" & searchbox_txt.Text & "%' and  std_t.s_no=place_rg_t.s_no ",
                                        rs_std_src.Open "select std_t.* from std_t , std_round_t where std_t.s_no like '" & searchbox_txt.Text & "%' and  std_t.s_no=std_round_t.s_no and rid=" & roundid & " and std_round_t.pid=" & att_pd_frm.pd_id & "  ", pms_cn
                                    End If
                             End If
                    Case 4
                             Set rs = New ADODB.Recordset
                            rs.CursorLocation = adUseClient
                              rs.Open "select max(rno) from place_round_t where pid=" & slec_std_frm.pd_id & "", pms_cn, 3, 3
                            MsgBox rs.Fields(0)
                             
                             Set rs_std_src = New ADODB.Recordset
                             rs_std_src.CursorLocation = adUseClient
                            
                            If flag = True Then
                                rs_std_src.Open "select std_t.* from std_t,std_round_t  where std_t.s_name like '%" & searchbox_txt.Text & "%' and std_t.s_no=std_round_t.s_no and std_round_t.pid=" & slec_std_frm.pd_id & " and std_round_t.rid=" & rs.Fields(0) & "", pms_cn
                            ElseIf flag = False Then
                                rs_std_src.Open "select std_t.* from std_t,std_round_t where std_t.s_name like '%" & searchbox_txt.Text & "%' and std_t.s_no=std_round_t.s_no and place_rg_t.pid=" & slec_std_frm.pd_id & " ", pms_cn
                            End If
                 End Select
                 
                Set std_search_dg.DataSource = rs_std_src
                std_search_dg.Columns(0).Caption = "Scholar Number"
                std_search_dg.Columns(1).Caption = "Student Name"
                std_search_dg.Columns(2).Caption = "Course"
                std_search_dg.Columns(3).Caption = "10th per"
                std_search_dg.Columns(4).Caption = "12th per"
                std_search_dg.Columns(5).Caption = "graduation per"
                std_search_dg.Columns(6).Caption = "post gradu per"
                std_search_dg.Columns(7).Caption = "Address"
                std_search_dg.Columns(8).Caption = "Mobile Number"
                std_search_dg.Columns(9).Caption = "Email ID"
        Else
           Set std_search_dg.DataSource = Nothing
        End If
        
End Sub

Private Sub searchbox_txt_KeyPress(KeyAscii As Integer)
    If flag = False Then
        Select Case KeyAscii
                Case 48 To 57 'numaric
                Case 8      'backspace
                Case Else
                  KeyAscii = 0
        End Select
    End If
    
End Sub





Private Sub std_search_dg_Click()
   Dim i As Integer
   If std_search_dg.Row <> -1 Then
             i = std_search_dg.Row
             std_search_dg.RowBookmark (i)
             stdschno = std_search_dg.Columns(0)
   
        Select Case cldform
            Case 1
              
                std_frm.state = 2
                   
                With std_frm
                    .schno_txt.Enabled = True
                    .std_nm_txt.Enabled = True
                    .co_cmb.Enabled = True
                    .stdmobnum_txt.Enabled = True
                    .stdemail_txt.Enabled = True
                    .ad_txt.Enabled = True
                    .s_cmd.Enabled = True
                    .clr_cmd.Enabled = True
                    .del_cmd.Enabled = True
                    .tenth_per_txt.Enabled = True
                    .twelth_per_txt.Enabled = True
                    .grad_txt.Enabled = True
                    
                    If std_search_dg.Columns(6) <> "" Then
                       .pgrad_txt.Enabled = True
                    Else
                       .pgrad_txt.Enabled = False
                    End If
                    
                    .schno_txt.Text = std_search_dg.Columns(0)
                    .std_nm_txt.Text = std_search_dg.Columns(1)
                    .tenth_per_txt.Text = std_search_dg.Columns(3)
                    .twelth_per_txt.Text = std_search_dg.Columns(4)
                    .grad_txt.Text = std_search_dg.Columns(5)
                    .pgrad_txt.Text = std_search_dg.Columns(6)
                    .stdmobnum_txt.Text = std_search_dg.Columns(8)
                    .stdemail_txt.Text = std_search_dg.Columns(9)
                    .ad_txt.Text = std_search_dg.Columns(7)
                    
                      Dim s As Long
                        For s = 0 To .co_cmb.ListCount - 1
                            If .co_cmb.ItemData(s) = std_search_dg.Columns(2) Then
                              .co_cmb.ListIndex = s
                              Exit For
                            End If
                        Next
                    
                    
                End With
            Case 2
                    rg_frm.std_co_txt = rg_frm.co_cmb.Text
                    rg_frm.std_name_txt = std_search_dg.Columns(1)
                    rg_frm.std_tenth_per_txt = std_search_dg.Columns(3)
                    rg_frm.std_twelth_per_txt = std_search_dg.Columns(4)
                    rg_frm.std_grad_txt = std_search_dg.Columns(5)
                    rg_frm.std_pgrad_txt = std_search_dg.Columns(6)
                    rg_frm.scl_no_txt = std_search_dg.Columns(0)
                    rg_frm.state = 1
                    'Set rs_std_src = New ADODB.Recordset
                    
                    'rs_std_src.Open
            Case 3
                    att_pd_frm.std_name_txt = std_search_dg.Columns(1)
                    att_pd_frm.std_id = std_search_dg.Columns(0)
                    att_pd_frm.std_co = std_search_dg.Columns(2)
                    
                    att_pd_frm.add_cmd.Enabled = True
                    
        End Select
        Unload Me
   End If
End Sub
