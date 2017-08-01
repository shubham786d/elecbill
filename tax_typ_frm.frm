VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form tax_typ_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tax Type"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "tax_typ_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton watt_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Per Kilowatt"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9120
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   10560
      Picture         =   "tax_typ_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.OptionButton perunit_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Per Unit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.OptionButton fix_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fixed"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.OptionButton per_opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Percentage"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox taxname_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      MaxLength       =   255
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      Picture         =   "tax_typ_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      Picture         =   "tax_typ_frm.frx":15B3D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   4080
      Picture         =   "tax_typ_frm.frx":1610A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   8880
      Picture         =   "tax_typ_frm.frx":16637
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Qty_msk 
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "000.00"
      Mask            =   "###.##"
      PromptChar      =   "0"
   End
   Begin VB.Label sidevalue_lbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TAX Name :"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label type_lbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      Height          =   6375
      Left            =   2880
      Top             =   600
      Width           =   8895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   2880
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "tax_typ_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tax_id As Long
Public state As Integer
Public oldtaxname As String
Dim rs_taxty As ADODB.Recordset
Private Sub Text2_Change()

End Sub

Private Sub Command1_Click()
'MsgBox Qty_msk.ClipText
'Text1.Text = Qty_msk.Text

End Sub

Private Sub del_cmd_Click()
If state = 2 Then
        Dim test As Integer
        test = MsgBox("Do U Want To Delete This Record ?", vbYesNoCancel + vbQuestion, "Information")
         
         If test = 6 Then
            Dim qry As String
            qry = "delete from taxtype_t where tid=" & taxtyp_src_frm.taxid_id & " "
            Call delete(qry)
                   
            MsgBox "Tax Type Record Deleted Successfully", vbInformation
          
            state = 3
            s_cmd.Enabled = False
            del_cmd.Enabled = False
            
            per_opt.value = True
            
            taxname_txt.Text = ""
            Qty_msk.Text = "000.00"
            
            taxname_txt.Enabled = False
            per_opt.Enabled = False
            perunit_opt.Enabled = False
            watt_opt.Enabled = False
            Qty_msk.Enabled = False
            fix_opt.Enabled = False
    End If
End If
End Sub

Private Sub exit_cmd_Click()
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub fix_opt_Click()
type_lbl.Caption = "Fixed TAX"
sidevalue_lbl.Caption = "Rs"
End Sub

Private Sub Form_Load()
    Set rs_taxty = New ADODB.Recordset
    rs_taxty.CursorLocation = adUseClient
       
    rs_taxty.Open ("select max(tid) from taxtype_t"), bms_cn, 3, 3
    
    If rs_taxty.RecordCount > 0 Then
         If IsNull(rs_taxty.Fields(0)) Then
             tax_id = 1
         Else
             tax_id = rs_taxty.Fields(0) + 1
         End If
    End If
    per_opt.value = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub new_cmd_Click()
s_cmd.Enabled = True
'clr_cmd.Enabled = True

taxname_txt.Enabled = True
per_opt.Enabled = True
perunit_opt.Enabled = True
Qty_msk.Enabled = True
fix_opt.Enabled = True
watt_opt.Enabled = True
del_cmd.Enabled = False
state = 1
End Sub

Private Sub per_opt_Click()
type_lbl.Caption = "Percentage TAX"
sidevalue_lbl.Caption = "%"
End Sub

Private Sub perunit_opt_Click()
type_lbl.Caption = "Per Unit TAX"
sidevalue_lbl.Caption = "Rs"
End Sub

Private Sub Qty_msk_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case 46
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub s_cmd_Click()
Select Case state
    Case 1
            If taxname_txt <> "" Then
                If Qty_msk.Text <> "000.00" Then
                    Call isduplicate
                    If isduplicate = False Then
                            Dim str As String
                            If per_opt.value = True Then
                                 str = "insert into taxtype_t values('" & tax_id & "','" & taxname_txt.Text & "','" & Qty_msk.Text & "',NULL,NULL,NULL)"
                            ElseIf fix_opt.value = True Then
                                 str = "insert into taxtype_t values('" & tax_id & "','" & taxname_txt.Text & "',NULL,'" & Qty_msk.Text & "',NULL,NULL)"
                            ElseIf perunit_opt.value = True Then
                                 str = "insert into taxtype_t values('" & tax_id & "','" & taxname_txt.Text & "',NULL,NULL,'" & Qty_msk.Text & "',NULL)"
                            ElseIf watt_opt.value = True Then
                                 str = "insert into taxtype_t values('" & tax_id & "','" & taxname_txt.Text & "',NULL,NULL,NULL,'" & Qty_msk.Text & "')"
                            End If
                            
                            
                            insert (str)
                            MsgBox "New Record Saved Successfully", vbInformation
                            
                            state = 3
                            tax_id = tax_id + 1
                            per_opt.value = True
                            
                            taxname_txt.Text = ""
                            Qty_msk.Text = "000.00"
                            
                            s_cmd.Enabled = False
                            'clr_cmd.Enabled = False
                            
                            taxname_txt.Enabled = False
                            per_opt.Enabled = False
                            perunit_opt.Enabled = False
                            watt_opt.Enabled = False
                            Qty_msk.Enabled = False
                            fix_opt.Enabled = False
                    Else
                             MsgBox "This TAX Name Is Already Used ", vbInformation
                    End If
                Else
                         MsgBox "Please Input Tax Values", vbInformation
                End If
            Else
                    MsgBox "Please Input TAX Name  ", vbInformation
            End If
    Case 2
        
            If isduplicate = False Then
                Dim qry As String
                If per_opt.value = True Then
                     qry = "update taxtype_t set tname='" & taxname_txt.Text & "' ,perc='" & Qty_msk.Text & "', fixed=NULL,perunit=NULL,perkilowatt=NULL where tid =" & taxtyp_src_frm.taxid_id & " "
                ElseIf fix_opt.value = True Then
                     qry = "update taxtype_t set tname='" & taxname_txt.Text & "' ,perc= Null,fixed='" & Qty_msk.Text & "',perunit=Null,perkilowatt=NULL where tid =" & taxtyp_src_frm.taxid_id & ""
                ElseIf perunit_opt.value = True Then
                     qry = "update taxtype_t set tname='" & taxname_txt.Text & "' ,perc=  Null  ,fixed=Null,perkilowatt=NULL,perunit='" & Qty_msk.Text & "' where tid =" & taxtyp_src_frm.taxid_id & ""
                ElseIf watt_opt.value = True Then
                     qry = "update taxtype_t set tname='" & taxname_txt.Text & "' ,perc=  Null  ,fixed=Null,perunit=NULL,perkilowatt='" & Qty_msk.Text & "' where tid =" & taxtyp_src_frm.taxid_id & ""
        
                End If
                
                
                Call update(qry)
                
                state = 3
                per_opt.value = True
                    
                taxname_txt.Text = ""
                Qty_msk.Text = "000.00"
                
                taxname_txt.Enabled = False
                per_opt.Enabled = False
                perunit_opt.Enabled = False
                watt_opt.Enabled = False
                Qty_msk.Enabled = False
                fix_opt.Enabled = False
                
                s_cmd.Enabled = False
                del_cmd.Enabled = False
                MsgBox "Record Update SuccessFully", vbInformation
            Else
                MsgBox "This TAX Name Is Already Used ", vbInformation
            End If
    Case 3
            
End Select
End Sub

Private Function isduplicate() As Boolean
    Set rs_taxty = New ADODB.Recordset
    rs_taxty.CursorLocation = adUseClient
    
    If oldtaxname <> taxname_txt.Text Then
        rs_taxty.Open ("select tname from taxtype_t where tname='" & taxname_txt.Text & "'"), bms_cn, 3, 3

        If rs_taxty.RecordCount > 0 Then
            isduplicate = True
        ElseIf rs_taxty.RecordCount = 0 Then
            isduplicate = False
        End If
    Else
        isduplicate = False
    End If
    
    
End Function

Private Sub src_cmd_Click()
taxtyp_src_frm.callvalue = 1
taxtyp_src_frm.Show vbModal
End Sub

Private Sub watt_opt_Click()
type_lbl.Caption = "Per Kilowatt Tax"
sidevalue_lbl.Caption = "Rs"
End Sub

