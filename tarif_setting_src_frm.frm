VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tarif_setting_src_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tarif_setting_src_frm.frx":0000
   ScaleHeight     =   7695
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   5880
      Picture         =   "tarif_setting_src_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
   End
   Begin VB.OptionButton type_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type"
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton purpose_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purpose"
      Height          =   495
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton phase_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phase"
      Height          =   495
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox text_txt 
      Height          =   375
      Left            =   4920
      MaxLength       =   255
      TabIndex        =   0
      Top             =   2400
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2655
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
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
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6735
      Left            =   1320
      Top             =   840
      Width           =   9855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   1320
      Top             =   720
      Width           =   9855
   End
End
Attribute VB_Name = "tarif_setting_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_src As ADODB.Recordset
'Dim tid As Long
'Dim pur_id As Long
'Dim pha_id As Long

Private Sub co_ex_cmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
type_opt.value = True
End Sub

Private Sub search_dg_Click()
        Dim i As Integer
        If search_dg.Row <> -1 Then
            i = search_dg.Row
            search_dg.RowBookmark (i)
            
            If type_opt.value = True Then
                tarif_setting_frm.update_id = search_dg.Columns(0)
                tarif_setting_frm.type_opt = True
                tarif_setting_frm.purpose_opt = False
                tarif_setting_frm.phase_opt = False
                
                tarif_setting_frm.type_opt.Enabled = True
                tarif_setting_frm.purpose_opt.Enabled = False
                tarif_setting_frm.phase_opt.Enabled = False
            ElseIf purpose_opt.value = True Then
                tarif_setting_frm.update_id = search_dg.Columns(0)
                tarif_setting_frm.type_opt = False
                tarif_setting_frm.purpose_opt = True
                tarif_setting_frm.phase_opt = False
                
                tarif_setting_frm.type_opt.Enabled = False
                tarif_setting_frm.purpose_opt.Enabled = True
                tarif_setting_frm.phase_opt.Enabled = False
                
            ElseIf phase_opt.value = True Then
                tarif_setting_frm.update_id = search_dg.Columns(0)
                tarif_setting_frm.type_opt = False
                tarif_setting_frm.purpose_opt = False
                tarif_setting_frm.phase_opt = True
                
                tarif_setting_frm.type_opt.Enabled = False
                tarif_setting_frm.purpose_opt.Enabled = False
                tarif_setting_frm.phase_opt.Enabled = True
            End If
            
            tarif_setting_frm.text_txt.Enabled = True
            tarif_setting_frm.s_cmd.Enabled = True
            tarif_setting_frm.del_cmd.Enabled = True
            tarif_setting_frm.state = 2
            
            tarif_setting_frm.text_txt.Text = search_dg.Columns(1)
            tarif_setting_frm.oldtaxname = search_dg.Columns(1)
            Unload Me
        End If
End Sub

'If type_opt.Value = True Then
'ElseIf purpose_opt.Value = True Then
'ElseIf phase_opt.Value = True Then
'End If

Private Sub type_opt_Click()
Label1.Caption = "Type Name:"
text_txt.Text = ""
End Sub
Private Sub phase_opt_Click()
Label1.Caption = "Phase Name:"
text_txt.Text = ""
End Sub

Private Sub purpose_opt_Click()
Label1.Caption = "Purpose Name:"
text_txt.Text = ""
End Sub
Private Sub text_txt_Change()
        If text_txt.Text <> "" Then
            Set rs_src = New ADODB.Recordset
            rs_src.CursorLocation = adUseClient

            If type_opt.value = True Then
               rs_src.Open "select * from type_t where tname like '%" & text_txt.Text & "%' ", bms_cn, 3, 3
               Set search_dg.DataSource = rs_src
               search_dg.Columns(1).Caption = "Type Name"
            ElseIf purpose_opt.value = True Then
               rs_src.Open "select * from perpose_t where pname like '%" & text_txt.Text & "%' ", bms_cn, 3, 3
               Set search_dg.DataSource = rs_src
               search_dg.Columns(1).Caption = "Perpose Name"
            ElseIf phase_opt.value = True Then
               rs_src.Open "select * from phase_t where pname like '%" & text_txt.Text & "%' ", bms_cn, 3, 3
               Set search_dg.DataSource = rs_src
               search_dg.Columns(1).Caption = "Phase Name"
            End If
            
             search_dg.Columns(0).Visible = True
             
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub
                
