VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ro_type_src_frm 
   Caption         =   "Form2"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   Picture         =   "ro_type_src_frm.frx":0000
   ScaleHeight     =   4560
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   3600
      Picture         =   "ro_type_src_frm.frx":9F77D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid rtype_search_dg 
      Height          =   2055
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
   Begin VB.TextBox src_txt 
      Height          =   285
      Left            =   3720
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rounds Name:"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   1575
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
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   4095
      Left            =   360
      Top             =   240
      Width           =   7935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   600
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "ro_type_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_rt_src As ADODB.Recordset
Public rtype_id As Long

Private Sub exit_cmd_Click()
    Unload Me
End Sub

Private Sub rtype_search_dg_Click()
        Dim i As Integer
        If rtype_search_dg.Row <> -1 Then
            i = rtype_search_dg.Row
            rtype_search_dg.RowBookmark (i)
            rtype_id = rtype_search_dg.Columns(0)
            
            With rs_rt_src
                ro_type_frm.ro_txt = .Fields(1)
                ro_type_frm.ro_desc_txt = .Fields(2)
            End With
            ro_type_frm.state = 2
            
            ro_type_frm.ro_txt.Enabled = True
            ro_type_frm.ro_desc_txt.Enabled = True
            ro_type_frm.s_cmd.Enabled = True
            ro_type_frm.clr_cmd.Enabled = True
            ro_type_frm.del_cmd.Enabled = True
            Unload Me
        End If
End Sub

Private Sub src_txt_Change()
        If src_txt.Text <> "" Then
            Set rs_rt_src = New ADODB.Recordset
            rs_rt_src.CursorLocation = adUseClient
            
            
             rs_rt_src.Open "select * from round_typ_t where rname like '%" & src_txt.Text & "%' ", pms_cn, 3, 3
              
                     
             Set rtype_search_dg.DataSource = rs_rt_src
             rtype_search_dg.Columns(0).Visible = True
             rtype_search_dg.Columns(1).Caption = "Round Name"
             rtype_search_dg.Columns(2).Caption = "Round Description"
        Else
                Set rtype_search_dg.DataSource = Nothing
                
        End If
End Sub
