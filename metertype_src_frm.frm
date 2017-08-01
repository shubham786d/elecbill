VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form metertype_src_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "metertype_src_frm.frx":0000
   ScaleHeight     =   6720
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox searchbox_txt 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton exit_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      Picture         =   "metertype_src_frm.frx":B4AE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Meter Type Name"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rent Price"
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   2055
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
      Left            =   6960
      TabIndex        =   1
      Top             =   1680
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2655
      Left            =   840
      TabIndex        =   0
      Top             =   2400
      Width           =   8535
      _ExtentX        =   15055
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
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   2415
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
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   6135
      Left            =   600
      Top             =   600
      Width           =   9015
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   360
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "metertype_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_meter As ADODB.Recordset
Public metertypid As Long
Public callvalue As Long


Private Sub exit_cmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Option1.value = True
End Sub

Private Sub search_dg_Click()
            If search_dg.Row <> -1 Then
            Dim i As Integer
            i = search_dg.Row
            search_dg.RowBookmark (i)
            metertypid = search_dg.Columns(0)
                
                Select Case callvalue
                    Case 1
                        With metertype_frm
                             .mname_txt.Text = search_dg.Columns(1)
                             .rent_txt.Text = search_dg.Columns(2)
                             .oldtext = search_dg.Columns(1)
                             Call metertype_frm.meter_readingset
                             For i = 1 To .digit_cmb.ListCount
                               If .digit_cmb.List(i + 1) = search_dg.Columns(3) Then
                                    .digit_cmb.ListIndex = i + 1
                                    Exit For
                               End If
                             Next
                             
                             If i - 1 = .digit_cmb.ListCount Then
                                .digit_cmb.AddItem search_dg.Columns(3)
                                 .digit_cmb.ListIndex = .digit_cmb.NewIndex
                             End If
                             .mname_txt.Enabled = True
                             .rent_txt.Enabled = True
                             .digit_cmb.Enabled = True
                             .s_cmd.Enabled = True
                             .del_cmd.Enabled = True
                
                             .state = 2
                        End With
                    Case 2
                       With meter_frm
                            .metertypeid = search_dg.Columns(0)
                            .metertype_txt.Text = search_dg.Columns(1)
                            .meterrent_txt.Text = search_dg.Columns(2)
                            .meterread_txt.MaxLength = search_dg.Columns(3)
                       End With
                End Select
                    Unload Me
                End If
End Sub

Private Sub searchbox_txt_Change()
        If searchbox_txt.Text <> "" Then
            Set rs_meter = New ADODB.Recordset
            rs_meter.CursorLocation = adUseClient
            Dim qry As String
             
             
             If Option1.value = True Then
              qry = "select * from metertyp_t where mname like '%" & searchbox_txt.Text & "%' "
             ElseIf Option2.value = True Then
              qry = "select * from metertyp_t where rentprice >= '" & searchbox_txt.Text & "' "
             End If
             
             rs_meter.Open qry, bms_cn, 3, 3
             Set search_dg.DataSource = rs_meter
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Caption = "Meter Type Name"
             search_dg.Columns(2).Caption = "Meter Rent Price"
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub
