VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form reader_src_frm 
   Caption         =   "Form3"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form3"
   ScaleHeight     =   5070
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox src_txt 
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   4560
      Picture         =   "reader_src_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2895
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Reader Name :"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   1215
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
      Left            =   4320
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "reader_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_reader_src As ADODB.Recordset
Public callvalue As Integer

Private Sub exit_cmd_Click()
Unload Me
End Sub

Private Sub search_dg_Click()
    If search_dg.Row <> -1 Then
        Dim i As Long
         i = search_dg.Row
        search_dg.RowBookmark (i)
        Select Case callvalue
            Case 1
                reader_frm.searchedrederid = search_dg.Columns(0)
                
                With reader_frm
                    .readerid_txt.Text = search_dg.Columns(0)
                    .readname_txt.Text = search_dg.Columns(1)
                    .mob_txt.Text = search_dg.Columns(2)
                    .address_txt.Text = search_dg.Columns(3)
                    .readname_txt.Enabled = True
                    .mob_txt.Enabled = True
                    .address_txt.Enabled = True
        
                    .del_cmd.Enabled = True
                    .s_cmd.Enabled = True
                    .state = 2
                End With
            Case 2
                With con_frm
                    .searchreader_id = search_dg.Columns(0)
                    .readname_txt.Text = search_dg.Columns(1)
                End With
        End Select
    
        Unload Me
    End If
End Sub

Private Sub src_txt_Change()
        If src_txt.Text <> "" Then
            Set rs_reader_src = New ADODB.Recordset
            rs_reader_src.CursorLocation = adUseClient
            Dim str As String
            'str = "Rid,RName,Mobileno,address"
             
                rs_reader_src.Open "select * from reader_t  where  rname like '%" & src_txt.Text & "%' ", bms_cn, 3, 3
             Set search_dg.DataSource = rs_reader_src
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Caption = "Round Name"
             search_dg.Columns(2).Caption = "Round Description"
             'search_dg.Columns(6).Visible = True
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub
