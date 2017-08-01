VERSION 5.00
Begin VB.Form metertype_frm 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meter Type Form"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "metertype_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox digit_cmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   10560
      Picture         =   "metertype_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      Picture         =   "metertype_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      Picture         =   "metertype_frm.frx":15B3D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   4800
      Picture         =   "metertype_frm.frx":1610A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   9240
      Picture         =   "metertype_frm.frx":16637
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox rent_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   1
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox mname_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      MaxLength       =   255
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Digits"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Type Name"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rent Price"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5415
      Left            =   3480
      Top             =   840
      Width           =   7935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   3480
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "metertype_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_metertype As ADODB.Recordset
Dim meter_id As Long
Public state As Integer
Public oldtext As String
Dim minmeterdigit As Integer


Private Sub co_ex_cmd_Click()
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub del_cmd_Click()
    If state = 2 Then
        Dim test As Integer
        test = MsgBox("Do U Want To Delete This Record ?", vbYesNoCancel + vbQuestion, "Information")
         
         If test = 6 Then
            Dim qry As String
            qry = "delete from metertyp_t where mtid=" & metertype_src_frm.metertypid & " "
            Call delete(qry)
                   
            MsgBox "Meter Type Record Deleted Successfully", vbInformation
          
            state = 3
            
            mname_txt.Enabled = False
            rent_txt.Enabled = False
            digit_cmb.Enabled = False
            
            mname_txt.Text = ""
            rent_txt.Text = ""
            ' = ""
            oldtext = ""
            
            s_cmd.Enabled = False
            'clr_cmd.Enabled = False
            del_cmd.Enabled = False
        End If
    End If
End Sub

Private Sub digit_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
End Select
End Sub

Private Sub Form_Load()
   minmeterdigit = 5
   
   Set rs_metertype = New ADODB.Recordset                           ' TYPE ID
    rs_metertype.CursorLocation = adUseClient
       
    rs_metertype.Open ("select max(mtid) from metertyp_t"), bms_cn, 3, 3
   ' rs_metertype
   
    If rs_metertype.RecordCount > 0 Then
         If IsNull(rs_metertype.Fields(0)) Then
             meter_id = 1
         Else
             meter_id = rs_metertype.Fields(0) + 1
         End If
    End If
    
    rs_metertype.Close

state = 3
    
    Call meter_readingset
    
End Sub
Public Sub meter_readingset()
    digit_cmb.Clear
    digit_cmb.AddItem "--- please select Number Of Digit ---"
    
    Set rs_metertype = New ADODB.Recordset                           ' meter digit len set
    rs_metertype.CursorLocation = adUseClient
    
    rs_metertype.Open ("select * from settings_t"), bms_cn, 3, 3
    Dim i As Integer
    For i = minmeterdigit To rs_metertype.Fields(7)
       digit_cmb.AddItem i
    Next
    
    digit_cmb.ListIndex = 0
    rs_metertype.Close
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
Call meter_readingset
s_cmd.Enabled = True

mname_txt.Enabled = True
rent_txt.Enabled = True
digit_cmb.Enabled = True
state = 1
End Sub

Private Sub rent_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
End Select
End Sub

'mname_txt
'rent_txt
Private Sub s_cmd_Click()
        
        Select Case state
            Case 1
                    If mname_txt <> "" Then
                        If rent_txt <> "" Then
                            If digit_cmb.ListIndex <> 0 Then
                                If isdup() = False Then
                                    
                                    Dim str As String
                                    str = "insert into metertyp_t values('" & meter_id & "','" & mname_txt.Text & "','" & rent_txt.Text & "','" & digit_cmb.Text & "')"
                                    Call insert(str)
                                    MsgBox "New Record Saved Successfully", vbInformation
                                    state = 3
                                    mname_txt.Enabled = False
                                    rent_txt.Enabled = False
                                    digit_cmb.Enabled = False
                                    
                                    mname_txt.Text = ""
                                    rent_txt.Text = ""
                                    oldtext = ""
                                    digit_cmb.ListIndex = 0
                                    meter_id = meter_id + 1
                                    
                                    s_cmd.Enabled = False
                                        
                                Else
                                    MsgBox "This type of meter is already existing", vbInformation
                                End If
                            Else
                                MsgBox "Please Input Number Of Digit Of Meter ", vbInformation
                            End If
                        Else
                            MsgBox "Please Input Meter Rent ", vbInformation
                        End If
                     Else
                        MsgBox "Please Input Meter Name ", vbInformation
                     End If
            Case 2
                 If mname_txt <> "" Then
                        If rent_txt <> "" Then
                            If digit_cmb.ListIndex <> 0 Then
                                 If (isdup() = False) Or mname_txt.Text = oldtext Then
                                    str = "update metertyp_t set mname='" & mname_txt.Text & "' , rentprice='" & rent_txt.Text & "',digits='" & digit_cmb.Text & "' where mtid=" & metertype_src_frm.metertypid & ""
                                    Call update(str)
                                    state = 3
                                    
                                    mname_txt.Enabled = False
                                    rent_txt.Enabled = False
                                    digit_cmb.Enabled = False
                                    
                                    mname_txt.Text = ""
                                    rent_txt.Text = ""
                                    digit_cmb.ListIndex = 0
                                    
                                    s_cmd.Enabled = False
                                    
                                    del_cmd.Enabled = False
                                    oldtext = ""
                                    MsgBox "Record Update SuccessFully", vbInformation
                                
                                Else
                                    MsgBox "This Name Is Already Used ", vbInformation
                                End If
                            Else
                                MsgBox "Please Input Number Of Digit Of Meter ", vbInformation
                            End If
                        Else
                            MsgBox "Please Input Meter Rent ", vbInformation
                        End If
                     Else
                        MsgBox "Please Input Meter Name ", vbInformation
                     End If
            Case 3
        
        End Select
        Call meter_readingset
End Sub

Private Sub src_cmd_Click()
metertype_src_frm.callvalue = 1
metertype_src_frm.Show vbModal
End Sub

Private Function isdup() As Boolean
    
    Set rs_metertype = New ADODB.Recordset
    rs_metertype.CursorLocation = adUseClient
    
    If oldtext <> mname_txt.Text Then
        
        rs_metertype.Open (" select * from metertyp_t where mname='" & mname_txt.Text & "'"), bms_cn, 3, 3

        If rs_metertype.RecordCount > 0 Then
            isdup = True
        ElseIf rs_metertype.RecordCount = 0 Then
            isdup = False
        End If
    Else
        isdup = True
    End If
End Function


Private Sub Text1_Change()
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
End Select
End Sub
