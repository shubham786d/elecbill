VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form con_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connection Form"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "cust_frm.frx":0000
   ScaleHeight     =   9180
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      TabIndex        =   61
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   7920
      TabIndex        =   43
      Top             =   600
      Width           =   5175
      Begin VB.OptionButton Option2 
         Caption         =   "Board Employee"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   55
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   54
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton meteridsrc_cmd 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4680
         TabIndex        =   53
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton reader_src_cmd 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   52
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox meternum_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   47
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox metertyp_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   46
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox meterstrtp_txt 
         Height          =   285
         Left            =   1920
         TabIndex        =   45
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox readname_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   44
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consumer Type"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Meter Serial Number"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Meter Type"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "meter starting reading"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "meter reader Name"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   2400
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   32
      Top             =   5160
      Width           =   10815
      Begin VB.TextBox load_txt 
         Height          =   285
         Left            =   4920
         TabIndex        =   37
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox secuamt_txt 
         Height          =   285
         Left            =   5160
         TabIndex        =   36
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox loadmin_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox samtmax 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox loadmax_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7800
         TabIndex        =   33
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load Sanctioned"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Security Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "<="
         Height          =   255
         Left            =   7560
         TabIndex        =   40
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "=>"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "=>"
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   2160
      TabIndex        =   13
      Top             =   600
      Width           =   5535
      Begin VB.CommandButton addaddress_cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox add_cmb 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1440
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton addcross_cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox cname_txt 
         Height          =   285
         Left            =   1920
         TabIndex        =   24
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox mob_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox ivrs_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox phnno_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton namesrc_cmd 
         BackColor       =   &H8000000D&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton namesrc_cancel_cmd 
         BackColor       =   &H008080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox phn_chk 
         BackColor       =   &H8000000E&
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox mob_chk 
         BackColor       =   &H8000000E&
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox emailid_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   15
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CheckBox email_chk 
         BackColor       =   &H8000000E&
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   255
      End
      Begin MSComCtl2.DTPicker condate_dtp 
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   43712515
         CurrentDate     =   42465
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IVRS Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Address"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "mobile Number"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "phone  Number"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "connection Date"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tarif "
      Height          =   1215
      Left            =   2160
      TabIndex        =   5
      Top             =   3720
      Width           =   10815
      Begin VB.TextBox ctype_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton tarifsrc_cmd 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox purpose_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox phase_txt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "connection type"
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purpose"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phase given"
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   13200
      Picture         =   "cust_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton exit_cmd 
      Height          =   375
      Left            =   9720
      Picture         =   "cust_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   5040
      Picture         =   "cust_frm.frx":15917
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      Picture         =   "cust_frm.frx":15E44
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      Picture         =   "cust_frm.frx":16411
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "connection Date"
      Height          =   255
      Left            =   8880
      TabIndex        =   56
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   6825
      Left            =   1560
      Top             =   240
      Width           =   12615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Height          =   7095
      Left            =   1680
      Top             =   120
      Width           =   12375
   End
End
Attribute VB_Name = "con_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'searched = full searched
Public ivrs_searchedyesno As Integer   ' searched=1 else -999

Public consumer_id As Long
Public searchedconsumer_id As Long

Dim cmd_con As ADODB.Command

Dim rs_con As ADODB.Recordset 'connection
Dim rs_cnsu As ADODB.Recordset 'consumer

Dim minmonth As String ' store date of  stating connection

Public tarif_id As Long

Private ivsr_id As Long
Public searchedIvrsid As Long

Public searchreader_id As Long
                    
Public oldmeterid As Long

Public state As Integer ' 1:insert 2:update
Dim cuntype As String

Private Sub resetid()
     
   Set rs_con = New ADODB.Recordset
    rs_con.CursorLocation = adUseClient
       
    rs_con.Open ("select max(ivrs) from connection_t"), bms_cn, 3, 3
   ' rs_con
    
    If rs_con.RecordCount > 0 Then
         If IsNull(rs_con.Fields(0)) Then
             ivsr_id = 1
         Else
             ivsr_id = rs_con.Fields(0) + 1
         End If
    End If
       
  ivrs_txt.Text = ivsr_id
    rs_con.Close

End Sub



Private Sub add_txt_Change()

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub addaddress_cmd_Click()
    
    add_cmb.Visible = False
    Text1.Visible = True
    'Text1.SetFocus
    addcross_cmd.Visible = True
End Sub

Private Sub addcross_cmd_Click()
    Text1.Visible = False
    addcross_cmd.Visible = False
     Text1.Text = ""
     add_cmb.Visible = True
End Sub

Private Sub meteridsrc_cmd_Click()
meter_src_frm.callvalue = 2
meter_src_frm.Show vbModal
End Sub



Private Sub del_cmd_Click()
If state = 2 Then
            Dim test As Integer
            test = MsgBox("Do U Want To Delete This Record ?", vbYesNoCancel + vbQuestion, "Information")
            
             If test = 6 Then
                Set cmd_std = New ADODB.Command
                cmd_std.CommandType = adCmdText
                cmd_std.CommandText = "delete from std_t where s_no=" & std_src_frm.stdschno & " "
                cmd_std.ActiveConnection = pms_cn
                cmd_std.Execute
                
               
                MsgBox "Record Deleted Successfully"
                state = 3
                schno_txt.Text = ""
                std_nm_txt.Text = ""
                co_cmb.ListIndex = 0
                stdmobnum_txt.Text = ""
                stdemail_txt.Text = ""
                ad_txt.Text = ""

                s_cmd.Enabled = False
                del_cmd.Enabled = False
                
                schno_txt.Enabled = False
                std_nm_txt.Enabled = False
                co_cmb.Enabled = False
                stdmobnum_txt.Enabled = False
                stdemail_txt.Enabled = False
                ad_txt.Enabled = False
                clr_cmd.Enabled = False
            End If
         Else
             MsgBox "Please Search And Select Student For Delete", vbInformation
        End If
End Sub
    emailid_txt.Text = ""
    
    If email_chk.value = 1 Then

Private Sub email_chk_Click()
    If email_chk.value = 1 Then
        emailid_txt.Enabled = True
    Else
        emailid_txt.Enabled = False
    End If
End Sub

Private Sub Form_Load()
        Set rs_cnsu = New ADODB.Recordset
         rs_cnsu.CursorLocation = adUseClient
            
         rs_cnsu.Open ("select max(cid) from consumer_t"), bms_cn, 3, 3
        ' rs_cnsu
         
         If rs_cnsu.RecordCount > 0 Then
              If IsNull(rs_cnsu.Fields(0)) Then
                  consumer_id = 1
              Else
                  consumer_id = rs_cnsu.Fields(0) + 1
              End If
         End If
         
         rs_cnsu.Close
        
        
        cname_txt.Enabled = False
        mob_txt.Enabled = False
        phnno_txt.Enabled = False
        add_cmb.Enabled = False
        ivrs_txt.Enabled = False
        meternum_txt.Enabled = False
        load_txt.Enabled = False
        condate_dtp.Enabled = False
        secuamt_txt.Enabled = False
        meterstrtp_txt.Enabled = False
        ctype_txt.Enabled = False
        
        namesrc_cmd.Enabled = False
        ivrs_searchedyesno = -999
      Call resetid
End Sub

Private Sub load_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub mob_chk_Click()
    If mob_chk = 1 Then
        mob_txt.Enabled = True
    Else
        mob_txt.Enabled = False
    End If
    
    mob_txt.Text = ""
End Sub

Private Sub mob_txt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub namesrc_cancel_cmd_Click()
ivrs_searchedyesno = -999
namesrc_cancel_cmd.Visible = False
cname_txt.Enabled = True
addaddress_cmd.Visible = False
addcross_cmd.Visible = False
add_cmb.Visible = False
Text1.Visible = True
End Sub

Private Sub namesrc_cmd_Click()
con_src_frm.cldform = 2
con_src_frm.Show vbModal

End Sub

Private Sub new_cmd_Click()
        ivrs_searchedyesno = -999
        Text1.Enabled = True
        Option2.Enabled = True
        Option1.Enabled = True
        Option1.value = True
        cname_txt.Enabled = True
        add_cmb.Enabled = True
        mob_chk.Enabled = True
        Call mob_chk_Click
        mob_chk.value = 1
        phn_chk.Enabled = True
        email_chk.Enabled = True
        
        load_txt.Enabled = True
        condate_dtp.Enabled = True
        secuamt_txt.Enabled = True

        cname_txt.Text = ""
        mob_txt.Text = ""
        phnno_txt.Text = ""
        emailid_txt.Text = ""
        
        namesrc_cmd.Enabled = True
        reader_src_cmd.Enabled = True
        
        meteridsrc_cmd.Enabled = True
        tarifsrc_cmd.Enabled = True
        
        del_cmd.Enabled = False
        s_cmd.Enabled = True
        state = 1
End Sub

Private Sub phn_chk_Click()
    phnno_txt.Text = ""
    
    If phn_chk.value = 1 Then
        phnno_txt.Enabled = True
    Else
        phnno_txt.Enabled = False
    End If
End Sub

Private Sub phnno_txt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select

End Sub

Private Sub reader_src_cmd_Click()
reader_src_frm.callvalue = 2
reader_src_frm.Show vbModal
End Sub

Private Sub s_cmd_Click()
    
    Select Case state
        Case 1 ' insert
             If cname_txt.Text <> "" Then
                If (add_cmb.Visible = True And add_cmb.ListIndex > 0) Or (Text1.Text <> "" And Text1.Visible = True) Then
                     If ((Val(load_txt) >= Val(loadmin_txt.Text)) And (Val(load_txt) <= Val(loadmax_txt.Text))) Then
                         If Val(secuamt_txt.Text) >= Val(samtmax.Text) Then
                              
                            '//checking for starting date of con
                            Set rs_con = New ADODB.Recordset
                            rs_con.CursorLocation = adUseClient
                               
                             rs_con.Open ("select * from settings_t"), bms_cn, 3, 3
                             minmonth = rs_con.Fields(0)
                              
                             If CDate(minmonth) > condate_dtp.value Then
                                MsgBox "Reading can not be taken from less than starting of  reading date set by ADMIN which is " & minmonth & ""
                                Exit Sub
                             End If
                              
                              
                              
                              If mob_chk.value = 0 And phn_chk.value = 0 Then
                                    MsgBox "Please Input Mobile Number Or Number ", vbInformation
                                    Exit Sub
                                End If
                                
                                If mob_chk.value = 1 Then
                                    If Len(mob_txt.Text) <> 0 Then
                                        If Len(mob_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Mobile number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please Input Mobile Number", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                If phn_chk.value = 1 Then
                                    If Len(phnno_txt.Text) <> 0 Then
                                        If Len(phnno_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Phone number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please input Phone number ", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                If email_chk.value = 1 Then
                                    If emailid_txt.Text = "" Then
                                        MsgBox "Please Input EMail Address", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                    
                                If meternum_txt.Text = "" Then
                                    MsgBox "Please Select meter Number", vbInformation
                                    Exit Sub
                                End If
                                
                                If readname_txt.Text = "" Then
                                    MsgBox "Please Select meter Reader Name", vbInformation
                                    Exit Sub
                                End If
                                
                                If ctype_txt.Text = "" Then
                                    MsgBox "Please Select Tarif Type", vbInformation
                                    Exit Sub
                                End If
                                
                                If load_txt.Text = "" Then
                                    MsgBox "Please Input Load amount", vbInformation
                                    Exit Sub
                                End If
                                
                                If secuamt_txt.Text = "" Then
                                    MsgBox "Please Input Security Amount", vbInformation
                                    Exit Sub
                                End If
                                
                                If Option1.value = True Then
                                    cuntype = "N"
                                Else
                                    cuntype = "B"
                                End If

                                bms_cn.BeginTrans
                                Dim qry As String
                                If ivrs_searchedyesno = 1 Then   ' searched name/ivrs
                                      qry = "update consumer_t set mobno='" & mob_txt.Text & "',phno='" & phnno_txt.Text & "',emailid='" & emailid_txt.Text & "' where cid =" & consumer_id & ""
                                      Call update(qry)           ' consumer update
                                ElseIf ivrs_searchedyesno = -999 Then
                                      qry = " insert into consumer_t values('" & consumer_id & "','" & cname_txt.Text & "','" & mob_txt.Text & "','" & phnno_txt.Text & "','" & emailid_txt.Text & "')"
                                      Call insert(qry)
                                End If
                                    Dim address As String
                                    
                                    If add_cmb.Visible = True Then
                                        address = add_cmb.ItemData(add_cmb.ListIndex)
                                    Else
                                        address = Text1.Text
                                    End If
                                    
                                    qry = "insert into connection_t values('" & consumer_id & "','" & ivsr_id & "','" & meternum_txt & "','" & tarif_id & "','" & load_txt & "','" & condate_dtp.value & "','" & secuamt_txt.Text & "','" & meterstrtp_txt.Text & "','" & 1 & "','" & searchreader_id & "','" & address & "','""','" & cuntype & "','0')"
                                    Call insert(qry)
                                    
                                    qry = "update meter_t set constatus=1  where mid=" & meternum_txt & " "
                                    Call update(qry)           'meter coection set
                                    
                                    bms_cn.CommitTrans
                                    
                                    consumer_id = consumer_id + 1
                                    state = 3
                                    ivrs_txt.Text = Val(ivrs_txt.Text) + 1
                                    Call form_reset
                                    
                                     MsgBox "New Record Saved Successfully", vbInformation
                                
                            Else
                                MsgBox "security amount should be >= " & samtmax.Text & " for this Tariff Type", vbInformation
                                secuamt_txt.SetFocus
                            End If
                        Else
                            MsgBox "This Trafic only can take Load between " & loadmin_txt.Text & "-" & loadmax_txt.Text & "", vbInformation
                        End If
                    Else
                        MsgBox "Please Input Cunsumer Address", vbInformation
                    End If
                Else
                    MsgBox "Please Input Cunsumer Name", vbInformation
                End If
        Case 2 'update
             If cname_txt.Text <> "" Then
                If (add_cmb.Visible = True And add_cmb.ListIndex > 0) Or (Text1.Text <> "" And Text1.Visible = True) Then
                     If ((Val(load_txt) >= Val(loadmin_txt.Text)) And (Val(load_txt) <= Val(loadmax_txt.Text))) Then
                         If Val(secuamt_txt.Text) >= Val(samtmax.Text) Then
                              
                              If mob_chk.value = 0 And phn_chk.value = 0 Then
                                    MsgBox "Please Input Mobile Number Or Number ", vbInformation
                                    Exit Sub
                                End If
                                
                                If mob_chk.value = 1 Then
                                    If Len(mob_txt.Text) <> 0 Then
                                        If Len(mob_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Mobile number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please Input Mobile Number", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                If phn_chk.value = 1 Then
                                    If Len(phnno_txt.Text) <> 0 Then
                                        If Len(phnno_txt.Text) = 10 Then
                                        Else
                                            MsgBox "Phone number can not be less than 10 digit", vbInformation
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox "Please input Phone number ", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                
                                If email_chk.value = 1 Then
                                    If emailid_txt.Text = "" Then
                                        MsgBox "Please Input EMail Address", vbInformation
                                        Exit Sub
                                    End If
                                End If
                                    
                                If meternum_txt.Text = "" Then
                                    MsgBox "Please Select meter Number", vbInformation
                                    Exit Sub
                                End If
                                
                                If readname_txt.Text = "" Then
                                    MsgBox "Please Select meter Reader Name", vbInformation
                                    Exit Sub
                                End If
                                
                                If ctype_txt.Text = "" Then
                                    MsgBox "Please Select Tarif Type", vbInformation
                                    Exit Sub
                                End If
                                
                                If load_txt.Text = "" Then
                                    MsgBox "Please Input Load amount", vbInformation
                                    Exit Sub
                                End If
                                
                                If secuamt_txt.Text = "" Then
                                    MsgBox "Please Input Security Amount", vbInformation
                                    Exit Sub
                                End If
                                
                                    
                                If add_cmb.Visible = True Then
                                    address = add_cmb.ItemData(add_cmb.ListIndex)
                                Else
                                    address = Text1.Text
                                End If
                                    
                                 qry = "update consumer_t set mobno='" & mob_txt.Text & "',phno='" & phnno_txt.Text & "',emailid='" & emailid_txt.Text & "' where cid =" & searchedconsumer_id & ""
                                 Call update(qry)           ' consumer update
                                
                                 qry = "update connection_t set cid=" & searchedconsumer_id & ",meter_id='" & meternum_txt & "',tarif_id='" & tarif_id & "',Load='" & load_txt & "',cdate='" & condate_dtp.value & "',secuamt='" & secuamt_txt.Text & "',mstartrrd='" & meterstrtp_txt.Text & "',readerid=" & searchreader_id & ",address='" & address & "' where ivrs='" & searchedIvrsid & "'"
                                 Call update(qry)
                                 
                                 If oldmeterid <> Val(meternum_txt) Then
                                    qry = "update meter_t set constatus=0  where mid=" & oldmeterid & " "
                                    Call update(qry)           'meter coection set
                                    
                                    qry = "update meter_t set constatus=1  where mid=" & meternum_txt & " "
                                    Call update(qry)           'meter coection set
                                 End If
                                 MsgBox "Record Updateted Successfully", vbInformation
                                 Call form_reset
                        Else
                            MsgBox "security amount should be >= " & samtmax.Text & " for this Tariff Type", vbInformation
                            secuamt_txt.SetFocus
                        End If
                    Else
                        MsgBox "This Trafic only can take Load between " & loadmin_txt.Text & "-" & loadmax_txt.Text & "", vbInformation
                    End If
                Else
                    MsgBox "Please Input Cunsumer Address", vbInformation
                End If
            Else
                MsgBox "Please Input Cunsumer Name", vbInformation
            End If
        Case 3
    End Select
        
End Sub

Private Sub secuamt_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub src_cmd_Click()
con_src_frm.cldform = 1
con_src_frm.Show vbModal
End Sub




'Private Function isdup() As Boolean
'
'          Dim i As Integer
'
'          For i = 0 To UBound(back_matrix)
'            If scl_no_txt.Text = back_matrix(i).sid And back_matrix(i).datastate <> 3 Then
'                isdup = True
'                Exit Function
'            End If
'          Next
'
'          isdup = False
'
'End Function
Private Sub tarifsrc_cmd_Click()
tarif_src_frm.callvalue = 3
 tarif_src_frm.Show vbModal
End Sub

Private Sub form_reset()
     state = 3
     s_cmd.Enabled = False
     del_cmd.Enabled = False
     
     condate_dtp.Enabled = False
     cname_txt.Enabled = False
     mob_txt.Enabled = False
     phnno_txt.Enabled = False
     add_cmb.Visible = False
     Text1.Visible = True
     Text1.Enabled = False
     emailid_txt.Enabled = False
     load_txt.Enabled = False
     secuamt_txt.Enabled = False
     
     meternum_txt.Text = ""
     metertyp_txt.Text = ""
     readname_txt.Text = ""

     ctype_txt.Text = ""
     phase_txt.Text = ""
     purpose_txt.Text = ""
     
     mob_chk.Enabled = False
     phn_chk.Enabled = False
     email_chk.Enabled = False
     
     mob_chk.value = False
     phn_chk.value = False
     email_chk.value = False
     
     cname_txt.Text = ""
     mob_txt.Text = ""
     phnno_txt.Text = ""
     Text1.Text = ""
     emailid_txt.Text = ""
     load_txt = ""
     secuamt_txt = ""
     loadmin_txt = ""
     loadmax_txt = ""
     samtmax = ""
     meterstrtp_txt = ""
     
     meteridsrc_cmd.Enabled = False
     namesrc_cmd.Enabled = False
     reader_src_cmd.Enabled = False
     tarifsrc_cmd.Enabled = False
     Option2.Enabled = False
     Option1.Enabled = False
        
     namesrc_cancel_cmd.Visible = False
     namesrc_cancel_cmd.Enabled = False
     
     addaddress_cmd.Visible = False
     addcross_cmd.Visible = False
     
     ivrs_searchedyesno = -999
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

