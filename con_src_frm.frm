VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form con_src_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "con_src_frm.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2655
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   12648447
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
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox searchbox_txt 
      Height          =   285
      Left            =   3840
      MaxLength       =   255
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consumer Name"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IVRS  Number"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton exit_cmd 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      Picture         =   "con_src_frm.frx":B4AE
      Style           =   1  'Graphical
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   720
      Top             =   600
      Width           =   9375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   480
      Top             =   360
      Width           =   9495
   End
End
Attribute VB_Name = "con_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public cldform As Integer  ' calling form value
Dim rs_src As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim rst2 As ADODB.Recordset
'Dim cmd_std_src As ADODB.Command
Dim flag As Boolean
Public stdschno As Long
'Dim rs_src As ADODB.Recordset

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
   ' search_dg.Col
End Sub

Private Sub Option1_Click()
    searchbox_txt.Visible = True
    clr_cmd.Visible = True
    searchLabel_lbl.Caption = "IVrs_src Number:"
    flag = False
    searchbox_txt.Text = ""
    searchbox_txt.MaxLength = 8
    searchbox_txt.SetFocus
End Sub

Private Sub Option2_Click()
    searchbox_txt.Visible = True
    clr_cmd.Visible = True
    searchLabel_lbl.Caption = "consumer Name:"
    flag = True
    searchbox_txt.Text = ""
    searchbox_txt.MaxLength = 255
    searchbox_txt.SetFocus
End Sub

Private Sub searchbox_txt_Change()
        If searchbox_txt.Text <> "" Then
              Set rs_src = New ADODB.Recordset
              rs_src.CursorLocation = adUseClient
                Select Case cldform
                    Case 1
                            Dim str As String
                           str = "ivrs,consumer_t.cid,cname,consumer_t.mobno,consumer_t.phno,consumer_t.emailid,meter_id,tarif_id,Load,cdate,secuamt,mstartrrd,address,readerid,cuntype"
                            If flag = True Then
                                rs_src.Open "select " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and consumer_t.cname like '%" & searchbox_txt.Text & "%' ", bms_cn
                            ElseIf flag = False Then
                                rs_src.Open "select " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and ivrs like '" & searchbox_txt.Text & "%' ", bms_cn
                            End If
                            
                    Case 2 ' con_frm
                            str = "ivrs,consumer_t.cid,cname,consumer_t.mobno,consumer_t.phno,consumer_t.emailid,meter_id,tarif_id,Load,cdate,secuamt,mstartrrd,readerid"
                            If flag = True Then
                                rs_src.Open "select distinct  " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and consumer_t.cname like '%" & searchbox_txt.Text & "%' ", bms_cn
                            ElseIf flag = False Then
                                rs_src.Open "select distinct " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and ivrs like '" & searchbox_txt.Text & "%' ", bms_cn
                            End If
                    Case 3 'reading frm
                            str = "ivrs,consumer_t.cid,cname,consumer_t.mobno,consumer_t.phno,consumer_t.emailid,meter_id,tarif_id,Load,cdate,secuamt,mstartrrd,address,readerid"
                            If flag = True Then
                                rs_src.Open "select " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and consumer_t.cname like '%" & searchbox_txt.Text & "%' ", bms_cn
                            ElseIf flag = False Then
                                rs_src.Open "select " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and ivrs like '" & searchbox_txt.Text & "%' ", bms_cn
                            End If
           
                    Case 4 'reading frm
                            str = "ivrs,consumer_t.cid,cname,consumer_t.mobno,consumer_t.phno,consumer_t.emailid,meter_id,tarif_id,Load,cdate,secuamt,mstartrrd,address,readerid"
                            If flag = True Then
                                rs_src.Open "select " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and consumer_t.cname like '%" & searchbox_txt.Text & "%' ", bms_cn
                            ElseIf flag = False Then
                                rs_src.Open "select " & str & " from consumer_t,connection_t where consumer_t.cid=connection_t.cid and ivrs like '" & searchbox_txt.Text & "%' ", bms_cn
                            End If
                            
                            
                            
                End Select
                
                Set search_dg.DataSource = rs_src
        Else
           Set search_dg.DataSource = Nothing
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





Private Sub search_dg_Click()
   Dim i As Integer
   If search_dg.Row <> -1 Then
             i = search_dg.Row
             search_dg.RowBookmark (i)
             con_frm.consumer_id = search_dg.Columns(0)
             
        Select Case cldform
            Case 1 ' con_frm update
                
                 With con_frm
                     .state = 2
                    
                     .searchedconsumer_id = search_dg.Columns(1)
                     .searchedIvrsid = search_dg.Columns(0)
                    .cname_txt.Text = search_dg.Columns(2)
                    
                    .condate_dtp = search_dg.Columns(9)
                    .load_txt.Text = search_dg.Columns(8)
                    .secuamt_txt = search_dg.Columns(10)
                    
                    '.add_cmb = search_dg.Columns(12)
                    
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT DISTINCT address from connection_t where cid=" & search_dg.Columns(1) & " ", bms_cn, 3, 3
                    Debug.Print "SELECT DISTINCT address from connection_t where cid=" & search_dg.Columns(1) & " "
                    
                    .add_cmb.Clear
                    .add_cmb.AddItem "--Select Address--"
                    
                    For i = 0 To rst.RecordCount - 1
                        .add_cmb.AddItem rst.Fields(0)
                        rst.MoveNext
                    Next
                    
                    For i = 1 To rst.RecordCount
                       If .add_cmb.List(i) = search_dg.Columns(12) Then
                            .add_cmb.ListIndex = i
                       End If
                    Next
                    
                
                    If search_dg.Columns(14) = "N" Then
                        .Option1.value = True
                    Else
                        .Option2.value = True
                    End If
                    
                    If search_dg.Columns(3) <> "" Then
                    .mob_chk.value = 1
                    .mob_txt.Text = search_dg.Columns(3)
                    End If
                    
                    If search_dg.Columns(4) <> "" Then
                        .phn_chk.value = 1
                        .phnno_txt.Text = search_dg.Columns(4)
                    End If
                    
                    If search_dg.Columns(5) <> "" Then
                        .email_chk.value = 1
                        .emailid_txt = search_dg.Columns(5)
                    End If
                    
                    Set rst = New ADODB.Recordset
                    rst.Open "select * from tarif_t where tarifid=" & search_dg.Columns(7) & "", bms_cn, 3, 3
                    
                    .loadmin_txt = rst.Fields("mload")
                    .loadmax_txt = rst.Fields(6)
                    .samtmax = rst.Fields(7)
                    
                    Set rst2 = New ADODB.Recordset
                    rst2.Open "select * from type_t where tid=" & rst.Fields(1) & "", bms_cn, 3, 3
                    .ctype_txt.Text = rst2.Fields(1)
                    
                     Set rst2 = New ADODB.Recordset
                     rst2.Open "select * from perpose_t where id=" & rst.Fields(1) & "", bms_cn, 3, 3
                    .purpose_txt.Text = rst2.Fields(1)
                    
                    Set rst2 = New ADODB.Recordset
                    rst2.Open "select * from phase_t where pid=" & rst.Fields(1) & "", bms_cn, 3, 3
                    .phase_txt.Text = rst2.Fields(1)
                    
                    .tarif_id = search_dg.Columns(7)
                    
                     Set rst = New ADODB.Recordset
                     
                    .oldmeterid = search_dg.Columns(6)
                    
                    rst.Open "select * from meter_t where mid=" & search_dg.Columns(6) & "", bms_cn, 3, 3
                    .meternum_txt.Text = search_dg.Columns(6)
                    .meterstrtp_txt.Text = rst.Fields(2)
                    
                    Set rst2 = New ADODB.Recordset
                    rst2.Open "select * from metertyp_t where mtid=" & rst.Fields(1) & "", bms_cn, 3, 3
                    .metertyp_txt.Text = rst2.Fields(1)
                    
                    Set rst = New ADODB.Recordset
                    rst.Open "select * from reader_t where rid=" & search_dg.Columns(13) & "", bms_cn, 3, 3
                    .readname_txt.Text = rst.Fields(1)
                    .searchreader_id = search_dg.Columns(7)
                    
                    
                    .add_cmb.Visible = True
                    .Text1.Visible = False
                    .addaddress_cmd.Visible = True
                    .Text1.Enabled = True
                    
                    .Option1.Enabled = True
                    .Option2.Enabled = True
                    
                    .mob_txt.Enabled = True
                    .phnno_txt.Enabled = True
                    .add_cmb.Enabled = True
                    
                    
                    .load_txt.Enabled = True
                    .condate_dtp.Enabled = True
                    .secuamt_txt.Enabled = True
                    
                     
                     .namesrc_cmd.Enabled = True
                    .reader_src_cmd.Enabled = True
                    .meteridsrc_cmd.Enabled = True
                    .tarifsrc_cmd.Enabled = True
                     
                     
                    .email_chk.Enabled = True
                    .mob_chk.Enabled = True
                    .phn_chk.Enabled = True
                    .s_cmd.Enabled = True
                End With
                 Unload Me
            Case 2 'con_frm insert
                With con_frm
                .consumer_id = search_dg.Columns(1)
                .cname_txt.Text = search_dg.Columns(2)
                
                Set rst = New ADODB.Recordset
                rst.Open "SELECT DISTINCT address from connection_t where cid=" & search_dg.Columns(1) & " ", bms_cn, 3, 3
                Debug.Print "SELECT DISTINCT address from connection_t where cid=" & search_dg.Columns(1) & " "
                
                .add_cmb.Clear
                .add_cmb.AddItem "--Select Address--"
                
                For i = 0 To rst.RecordCount - 1
                    .add_cmb.AddItem rst.Fields(0)
                    rst.MoveNext
                Next
                
                .add_cmb.ListIndex = 0
                
                If search_dg.Columns(3) <> "" Then
                .mob_txt.Text = search_dg.Columns(3)
                End If
                
                If search_dg.Columns(4) <> "" Then
                    .phn_chk.value = 1
                    .phnno_txt.Text = search_dg.Columns(4)
                End If
                
                If search_dg.Columns(5) <> "" Then
                    .email_chk.value = 1
                    .emailid_txt = search_dg.Columns(5)
                End If
                .cname_txt.Enabled = False
                If .state <> 2 Then
                    .namesrc_cancel_cmd.Visible = True
                    .ivrs_searchedyesno = 1
                    .add_cmb.Visible = True
                    .Text1.Visible = False
                    .addaddress_cmd.Visible = True
                End If
                
                End With
                Unload Me
            Case 3 'reading frm save
                With reading_frm
                If (CDate(search_dg.Columns(9)) < .readingofmonth_dtp.value) Then
                    'MsgBox .readingofmonth_dtp.value
                
                    .ivrs_txt.Text = search_dg.Columns(0)
                    .name_txt.Text = search_dg.Columns(2)
                    Set rst = New ADODB.Recordset
                    rst.Open "select * from  meter_t where mid =" & search_dg.Columns(6) & "", bms_cn, 3, 3
                    
                     .creading_txt = ""
                     .Check1.Visible = True
                     .Check1.value = 0
                     .Label2.Visible = True
                     
                    
                    .ivrs = search_dg.Columns(1)
                Else
                    MsgBox "Selected Customer's Connection Date can not be >  than reading Date | " & .readingofmonth_dtp.value
                    Exit Sub
                End If
                End With
                Unload Me
             Case 4 ' showbill
             
                showbill_frm.ivrs_txt = search_dg.Columns(0)
                Unload Me
        End Select
        
   End If
End Sub
