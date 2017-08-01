VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cnsu_src_frm 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "cnsu_src_frm.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2055
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
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
   Begin VB.TextBox src_txt 
      Height          =   285
      Left            =   3360
      MaxLength       =   255
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   3360
      Picture         =   "cnsu_src_frm.frx":B4AE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "             Name:"
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
      TabIndex        =   3
      Top             =   1080
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
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   240
      Top             =   240
      Width           =   7815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   4335
      Left            =   240
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "cnsu_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_cnsu_src As ADODB.Recordset
Public cnsu_id As Long


Private Sub exit_cmd_Click()
Unload Me
End Sub

Private Sub search_dg_Click()
        Dim i As Integer
        If search_dg.Row <> -1 Then
            i = search_dg.Row
            search_dg.RowBookmark (i)
            cnsu_id = search_dg.Columns(0)
            
            With cnsu_frm
            .cname_txt.Text = search_dg.Columns(1)
            
            If search_dg.Columns(2) <> "" Then
            .mob_chk.value = 1
            Call .mob_chk_Click
            .mob_txt.Text = search_dg.Columns(2)
            End If
            
            If search_dg.Columns(3) <> "" Then
                .phn_chk.value = 1
                Call .phn_chk_Click
                .phnno_txt.Text = search_dg.Columns(3)
            End If
            '.add_txt.Text = search_dg.Columns(4)
            
            If search_dg.Columns(4) <> "" Then
                .email_chk.value = 1
                Call .email_chk_Click
                .emailid_txt.Text = search_dg.Columns(4)
            End If

            .s_cmd.Enabled = True
            '.del_cmd.Enabled = True
            
            .cname_txt.Enabled = True
            '.mob_txt.Enabled = True
            '.phnno_txt.Enabled = True
            '.add_txt.Enabled = True
            
            .mob_chk.Enabled = True
            .phn_chk.Enabled = True
            .email_chk.Enabled = True
            End With
            cnsu_frm.state = 2
            
            
            Unload Me
        End If
End Sub

Private Sub src_txt_Change()
If src_txt.Text <> "" Then
            Set rs_cnsu_src = New ADODB.Recordset
            rs_cnsu_src.CursorLocation = adUseClient
            
            
             rs_cnsu_src.Open "select * from consumer_t where cname like '%" & src_txt.Text & "%' ", bms_cn, 3, 3
              
                     
             Set search_dg.DataSource = rs_cnsu_src
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Caption = "Round Name"
             search_dg.Columns(2).Caption = "Round Description"
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub
