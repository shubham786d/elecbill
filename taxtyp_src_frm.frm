VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form taxtyp_src_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "taxtyp_src_frm.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   3240
      Picture         =   "taxtyp_src_frm.frx":B4AE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox src_txt 
      Height          =   285
      Left            =   3360
      MaxLength       =   255
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2055
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   0
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
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tax Name:"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   1695
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
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   4095
      Left            =   480
      Top             =   360
      Width           =   8055
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0080C0FF&
      Height          =   4335
      Left            =   600
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "taxtyp_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_taxtyp_src As ADODB.Recordset
Public taxid_id As Long
Public callvalue As Integer

Private Sub exit_cmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rs_taxtyp_src = New ADODB.Recordset
            rs_taxtyp_src.CursorLocation = adUseClient
            
            
             rs_taxtyp_src.Open "select * from taxtype_t", bms_cn, 3, 3
              
                     
             Set search_dg.DataSource = rs_taxtyp_src
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Caption = "Round Name"
             search_dg.Columns(2).Caption = "Round Description"
End Sub

Private Sub search_dg_Click()
Dim i As Integer
        If search_dg.Row <> -1 Then
            
            
            i = search_dg.Row
            search_dg.RowBookmark (i)
            taxid_id = search_dg.Columns(0)
            
           Select Case callvalue
                Case 1  'tax_typ_frm

                    With tax_typ_frm
                    
                    .taxname_txt.Text = search_dg.Columns(1)
                    .oldtaxname = search_dg.Columns(1)
                    
                    If search_dg.Columns(2) <> "" Then
                        .per_opt.value = True
                        .Qty_msk.Text = search_dg.Columns(2)
                    ElseIf search_dg.Columns(3) <> "" Then
                        .fix_opt.value = True
                        .Qty_msk.Text = search_dg.Columns(3)
                    ElseIf search_dg.Columns(4) <> "" Then
                        .perunit_opt.value = True
                        .Qty_msk.Text = search_dg.Columns(4)
                    ElseIf search_dg.Columns(5) <> "" Then
                         .watt_opt.value = True
                         .Qty_msk.Text = search_dg.Columns(5)
                    End If
                    .s_cmd.Enabled = True
                    .del_cmd.Enabled = True
                    
                    .taxname_txt.Enabled = True
                    .per_opt.Enabled = True
                    .perunit_opt.Enabled = True
                    .Qty_msk.Enabled = True
                    .fix_opt.Enabled = True
                    .watt_opt.Enabled = True
                    .state = 2
                    End With
            Case 2 'tariftax_frm
                    With tariftax_frm
                    .taxname_txt.Text = search_dg.Columns(1)
                    
                    If search_dg.Columns(2) <> "" Then
                        .taxtype_txt = "Percentage"
                        .taxvalue_txt = search_dg.Columns(2)
                    ElseIf search_dg.Columns(3) <> "" Then
                        .taxtype_txt = "Fixed"
                        .taxvalue_txt = search_dg.Columns(3)
                    ElseIf search_dg.Columns(4) <> "" Then
                        .taxtype_txt = "Per Unit"
                        .taxvalue_txt = search_dg.Columns(4)
                    ElseIf search_dg.Columns(5) <> "" Then
                         .taxtype_txt = "Per Kilowatt"
                         .taxvalue_txt = search_dg.Columns(5)
                    End If
                    
                    .taxname_txt.Text = search_dg.Columns(1)
                    .taxid = search_dg.Columns(0)
                    End With
            
        
        End Select
            Unload Me
        End If

End Sub

Private Sub src_txt_Change()
        If src_txt.Text <> "" Then

            Set rs_taxtyp_src = New ADODB.Recordset
            rs_taxtyp_src.CursorLocation = adUseClient
            
            
             rs_taxtyp_src.Open "select * from taxtype_t where tname like '%" & src_txt.Text & "%' ", bms_cn, 3, 3
              
                     
             Set search_dg.DataSource = rs_taxtyp_src
             search_dg.Columns(0).Visible = True
             search_dg.Columns(1).Caption = "Round Name"
             search_dg.Columns(2).Caption = "Round Description"
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub
