VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form meter_src_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   2760
      Picture         =   "meter_src_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox src_txt 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5106
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
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter ID"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "meter_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_meter_src As ADODB.Recordset
Dim rst As ADODB.Recordset
Public callvalue As Integer

Private Sub Label8_Click()

End Sub

Private Sub search_dg_Click()
        If search_dg.Row <> -1 Then
            Dim i As Integer
            Select Case callvalue
                Case 1 'meter_frm
                        i = search_dg.Row
                        search_dg.RowBookmark (i)
                        meter_frm.searchedmeterID = search_dg.Columns(0)
                        
                        With meter_frm
                            .meterid_txt.Text = search_dg.Columns(0)
                            .metertype_txt.Text = search_dg.Columns(1)
                            .meterrent_txt = search_dg.Columns(2)
                            
                            If search_dg.Columns(4) = -1 Then
                                .yes_opt.value = True
                                .meterread_txt.Text = search_dg.Columns(5)
                                Set rst = New ADODB.Recordset
                                rst.CursorLocation = adUseClient
                                rst.Open "select * from metertyp_t where mtid=" & search_dg.Columns(6) & "", bms_cn, 3, 3
                                
                                .meterread_txt.MaxLength = rst.Fields(3)
                                
                            Else
                                .no_opt.value = True
                            End If
                            
                            .yes_opt.Enabled = True
                            .no_opt.Enabled = True
                            .metertypsrc_cmd.Enabled = True
                            .s_cmd.Enabled = True
                            .del_cmd.Enabled = True
                            .state = 2
                        End With
                Case 2 'con_frm
                        i = search_dg.Row
                        search_dg.RowBookmark (i)
                        'meterid = search_dg.Columns(0)
                        
                        With con_frm
                            .meternum_txt = search_dg.Columns(0)
                            .metertyp_txt = search_dg.Columns(1)
                            .meterstrtp_txt.Text = search_dg.Columns(5)
                        End With
                
            End Select
        
            Unload Me
        End If
End Sub

Private Sub src_txt_Change()
If src_txt.Text <> "" Then
            Set rs_meter_src = New ADODB.Recordset
            rs_meter_src.CursorLocation = adUseClient
            Dim str As String
            str = "mid,mname,rentprice,constatus,workstate,mstartread,mtypeid"
             If callvalue <> 2 Then
                rs_meter_src.Open "select " & str & " from meter_t,metertyp_t  where metertyp_t.mtid=meter_t.mtypeid and mid like '%" & src_txt.Text & "%' ", bms_cn, 3, 3
             Else
                rs_meter_src.Open "select " & str & " from meter_t,metertyp_t  where metertyp_t.mtid=meter_t.mtypeid and constatus <> -1 and mid like '%" & src_txt.Text & "%' ", bms_cn, 3, 3
             End If
             Set search_dg.DataSource = rs_meter_src
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Caption = "Round Name"
             search_dg.Columns(2).Caption = "Round Description"
             search_dg.Columns(6).Visible = True
        Else
             Set search_dg.DataSource = Nothing
        End If
End Sub

