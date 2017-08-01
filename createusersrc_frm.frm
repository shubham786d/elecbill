VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form createusersrc_frm 
   Caption         =   "Form3"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8640
   Picture         =   "createusersrc_frm.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   3600
      Picture         =   "createusersrc_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox src_txt 
      Height          =   285
      Left            =   3000
      MaxLength       =   255
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2055
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "           User ID"
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
      Top             =   1800
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
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   4455
      Left            =   480
      Top             =   480
      Width           =   7815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   4695
      Left            =   480
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "createusersrc_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As ADODB.Recordset

Private Sub search_dg_Click()
    
    Dim i As Integer
        If search_dg.Row <> -1 Then
            i = search_dg.Row
            search_dg.RowBookmark (i)
            With createuser_frm
            .username_txt = search_dg.Columns(0)
            .userId_txt = search_dg.Columns(1)
            .password_txt = search_dg.Columns(2)
            .oldname = search_dg.Columns(1)
            .DTPicker1 = search_dg.Columns(6)
            .DTPicker2 = search_dg.Columns(7)
            If search_dg.Columns(3) = 4 Then
                .admin_opt = True
            Else
                .normalusr_opt = True
                .normalusr_opt_Click
                If search_dg.Columns(3) = 3 Then
                    .bothrights_opt = True
                ElseIf search_dg.Columns(3) = 2 Then
                    .Tranrights_opt = True
                ElseIf search_dg.Columns(3) = 1 Then
                    .masterright_opt = True
                End If
            End If
            
            
            .username_txt.Enabled = True
            .userId_txt.Enabled = True
            .password_txt.Enabled = True
            .seepass_cmd.Enabled = True
            .admin_opt.Enabled = True
            .normalusr_opt.Enabled = True
            .masterright_opt.Enabled = True
            .Tranrights_opt.Enabled = True
            .bothrights_opt.Enabled = True
            .del_cmd.Enabled = True
            .DTPicker1.Enabled = True
            .DTPicker2.Enabled = True
            .s_cmd.Enabled = True
            .del_cmd.Enabled = True
            .state = 2
            Unload Me
            End With
        End If
        
End Sub

Private Sub src_txt_Change()
    If src_txt.Text <> "" Then
            Set rst = New ADODB.Recordset
            rst.CursorLocation = adUseClient
            
            
             rst.Open "select * from login_t where userid like '%" & src_txt.Text & "%' ", bms_cn, 3, 3
              
                     
             Set search_dg.DataSource = rst
             'search_dg.Columns(0).Visible = True
             search_dg.Columns(0).Caption = "Name"
             search_dg.Columns(1).Caption = " userid"
             search_dg.Columns(2).Caption = " Password"
             search_dg.Columns(3).Caption = " Rights"
             search_dg.Columns(4).Caption = " User type"
        Else
                Set search_dg.DataSource = Nothing
        End If
End Sub
