VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tarif_src_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10275
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tarif_src_frm.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ctype_cmb 
      Height          =   315
      Left            =   7560
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox phase_cmb 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox purpose_cmb 
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   5040
      Picture         =   "tarif_src_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2055
      Left            =   2400
      TabIndex        =   1
      Top             =   3360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3625
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH COURSE"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "connection type"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purpose"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phase given"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   6375
      Left            =   1320
      Top             =   480
      Width           =   8895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   1320
      Top             =   360
      Width           =   8895
   End
End
Attribute VB_Name = "tarif_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_tarif_src As ADODB.Recordset
Public tarifid As Long
Dim rs As New ADODB.Recordset
Public callvalue As Integer

Private Sub ctype_cmb_Click()
 If ctype_cmb.ListIndex <> -1 And purpose_cmb.ListIndex <> -1 And phase_cmb.ListIndex <> -1 Then

    Call src_txt_Change
End If
End Sub

Private Sub exit_cmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
    'Call src_txt_Change
    
    Dim query As String
    
    query = "select * from perpose_t"
    
    Call setcombo(query, purpose_cmb, "--Select Purpose--", 1, 0)
    
    query = "select * from phase_t"
    
    Call setcombo(query, phase_cmb, "--Select Phase--", 1, 0)
    
    query = "select * from type_t"
    
    Call setcombo(query, ctype_cmb, "--Select Type--", 1, 0)
    
End Sub

Private Sub src_txt_Change()
        If ctype_cmb.ListIndex <> 0 Or purpose_cmb.ListIndex <> 0 Or phase_cmb.ListIndex <> 0 Then
            Set rs_tarif_src = New ADODB.Recordset
            rs_tarif_src.CursorLocation = adUseClient
            Dim str As String
            str = "tarifid,typeid,type_t.tname,perposid,perpose_t.pname,phaseid,phase_t.pname,mmcprice,mload,maxload,minsecamt"
            'typeid perposid phaseid
             If ctype_cmb.ListIndex <> 0 And purpose_cmb.ListIndex = 0 And phase_cmb.ListIndex = 0 Then
                
                Debug.Print "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposeid=perpose_t.pid and tarif_t.phaseid=phase_t.pid  and typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & " "
                
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & " ", bms_cn, 3, 3
             ElseIf ctype_cmb.ListIndex = 0 And purpose_cmb.ListIndex <> 0 And phase_cmb.ListIndex = 0 Then
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and perposid=" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & " ", bms_cn, 3, 3
             
             ElseIf ctype_cmb.ListIndex = 0 And purpose_cmb.ListIndex = 0 And phase_cmb.ListIndex <> 0 Then
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and phaseid= " & phase_cmb.ItemData(phase_cmb.ListIndex) & " ", bms_cn, 3, 3
             
             ElseIf ctype_cmb.ListIndex <> 0 And purpose_cmb.ListIndex <> 0 And phase_cmb.ListIndex = 0 Then
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & " and perposid=" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & "", bms_cn, 3, 3
             
             ElseIf ctype_cmb.ListIndex = 0 And purpose_cmb.ListIndex <> 0 And phase_cmb.ListIndex <> 0 Then
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and perposid=" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & " and phaseid= " & phase_cmb.ItemData(phase_cmb.ListIndex) & " ", bms_cn, 3, 3
             
             ElseIf ctype_cmb.ListIndex <> 0 And purpose_cmb.ListIndex = 0 And phase_cmb.ListIndex <> 0 Then
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & "  and phaseid=" & phase_cmb.ItemData(phase_cmb.ListIndex) & " ", bms_cn, 3, 3
             
             ElseIf ctype_cmb.ListIndex <> 0 And purpose_cmb.ListIndex <> 0 And phase_cmb.ListIndex <> 0 Then
                rs_tarif_src.Open "select " & str & " from tarif_t,type_t,perpose_t,phase_t where tarif_t.typeid=type_t.tid and tarif_t.perposid=perpose_t.id and tarif_t.phaseid=phase_t.pid and typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & " and perposid=" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & " and phaseid=" & phase_cmb.ItemData(phase_cmb.ListIndex) & "", bms_cn, 3, 3
             End If
             
             Set search_dg.DataSource = rs_tarif_src
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Visible = False
             search_dg.Columns(3).Visible = False
             search_dg.Columns(5).Visible = False
             
             
             search_dg.Columns(2).Caption = "Connection type"
             search_dg.Columns(6).Caption = "Phase"
             search_dg.Columns(4).Caption = "Perpose"
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub



Private Sub phase_cmb_Click()
     If ctype_cmb.ListIndex <> -1 And purpose_cmb.ListIndex <> -1 And phase_cmb.ListIndex <> -1 Then
        Call src_txt_Change
     End If
End Sub

Private Sub purpose_cmb_Click()
     If ctype_cmb.ListIndex <> -1 And purpose_cmb.ListIndex <> -1 And phase_cmb.ListIndex <> -1 Then
         Call src_txt_Change
    End If
End Sub

Private Sub search_dg_Click()
        Dim i As Integer
        If search_dg.Row <> -1 Then
            i = search_dg.Row
            search_dg.RowBookmark (i)
            tarifid = search_dg.Columns(0)
            
            Select Case callvalue
                Case 1 'tarif_frm
            
                    With tarif_frm
                    .searched = True
                    Call updatecombo(.ctype_cmb, search_dg.Columns(1))
                    Call updatecombo(.phase_cmb, search_dg.Columns(5))
                    Call updatecombo(.purpose_cmb, search_dg.Columns(3))
                    'store value of cmbs
                    Call .sethistoryvalueofcmb(search_dg.Columns(1), search_dg.Columns(5), search_dg.Columns(3))
                    
                    
                    .mmc_txt = search_dg.Columns(7)
                    .minload_txt.Text = search_dg.Columns(8)
                    .maxload_txt.Text = search_dg.Columns(9)
                    .minamt_txt.Text = search_dg.Columns(10)
                    
                    Set rs = New ADODB.Recordset         ' flex grid set
                    rs.CursorLocation = adUseClient
                    rs.Open ("select * from tarifsetting_t where tarifid=" & search_dg.Columns(0) & "  "), bms_cn, 3, 3
                    If (rs.BOF <> True) Then
                    rs.MoveFirst
                    End If
                    
                    Call .update_backmatrix(-999, 0, 0, 0, 0)
                    .tarifgrid.Rows = 1
                    
                    For i = 1 To rs.RecordCount          ' back grid set
                         .tarifgrid.Rows = .tarifgrid.Rows + 1
                         .tarifgrid.TextMatrix(i, 0) = rs.Fields(2)
                         .tarifgrid.TextMatrix(i, 1) = rs.Fields(3)
                         .tarifgrid.TextMatrix(i, 2) = rs.Fields(4)
                         .tarifgrid.TextMatrix(i, 3) = rs.Fields(1)
                         If i = rs.RecordCount Then
                            .range1_txt = rs.Fields(3) + 1
                         End If
                         
                          Call .update_backmatrix(0, rs.Fields(2), rs.Fields(3), rs.Fields(1), rs.Fields(4))
                          
                         .backm_count = .backm_count + 1
                                      
                          rs.MoveNext
                     Next
                    
                    .fg_rowcount = rs.RecordCount
                    
                    .state = 2
                    .add_cmd.Enabled = True
                    .rem_cmd.Enabled = True
                    .rem_all_txt.Enabled = True
                    .s_cmd.Enabled = True
                    .phase_cmb.Enabled = True
                    .purpose_cmb.Enabled = True
                    .ctype_cmb.Enabled = True
                    .range2_txt.Enabled = True
                    .unitrate_txt.Enabled = True
                    .minload_txt.Enabled = True
                    .maxload_txt.Enabled = True
                    .minamt_txt.Enabled = True
                    .s_cmd.Enabled = True
                    .del_cmd.Enabled = True
                    
                    .Command1.Enabled = True
                    .tarifgrid.Enabled = True
                    .fix_cmd.Enabled = False
                    .mmc_txt.Enabled = False
                    End With
                Case 2 'tariftax_frm
                    Set rs = New ADODB.Recordset
                    rs.Open "select * from tariftax_t where tarif_id=" & search_dg.Columns(0) & "", bms_cn, 3, 3
                    
                    If rs.RecordCount = 0 Then
                        With tariftax_frm
                        .tarif_id = search_dg.Columns(0)
                        .ctype_txt.Text = search_dg.Columns(2)
                        .phase_txt.Text = search_dg.Columns(6)
                        .purpose_txt.Text = search_dg.Columns(4)
                        End With
                    Else
                        MsgBox "Tax's of selected tariff Type is already exisiting , Please search for exisiting Records", vbInformation
                    End If
                Case 3 'con_frm
                    With con_frm
                        .samtmax.Text = search_dg.Columns(10)
                       ' .maxloadamt = search_dg.Columns(9)
                        .loadmax_txt.Text = search_dg.Columns(9)
                       ' .minloadamt = search_dg.Columns(8)
                        .loadmin_txt = search_dg.Columns(8)
                        .tarif_id = search_dg.Columns(0)
                        .ctype_txt.Text = search_dg.Columns(2)
                        .phase_txt.Text = search_dg.Columns(6)
                        .purpose_txt.Text = search_dg.Columns(4)
                        
                        '.searchedtarifid = search_dg.Columns(0)
                        
                    End With
            End Select
            
                Unload Me
            
            
        End If
End Sub


Public Function updatestatefg(rid As Long)
    Static var As Integer
    
    If rid = -999 Then                     ' For Reseting all value of back grid
      ReDim fg_UState(0)
      var = 0
      Exit Function
    End If
      
   ' MsgBox var
    ReDim Preserve fg_UState(var)
    fg_UState(var).rid = rid
    fg_UState(var).datastate = 0
   
    Debug.Print fg_UState(var).rid; fg_UState(var).datastate
     var = var + 1
End Function
