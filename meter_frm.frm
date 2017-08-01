VERSION 5.00
Begin VB.Form meter_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meter Form"
   ClientHeight    =   7545
   ClientLeft      =   -2895
   ClientTop       =   3390
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "meter_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox meterread_txt 
      Height          =   285
      Left            =   7080
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   10320
      Picture         =   "meter_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox meterid_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   12
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   9000
      Picture         =   "meter_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   4800
      Picture         =   "meter_frm.frx":15917
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      Picture         =   "meter_frm.frx":15E44
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      Picture         =   "meter_frm.frx":16411
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.OptionButton no_opt 
      BackColor       =   &H8000000E&
      Caption         =   "No"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   4080
      Width           =   855
   End
   Begin VB.OptionButton yes_opt 
      BackColor       =   &H8000000E&
      Caption         =   "Yes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton metertypsrc_cmd 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9480
      TabIndex        =   2
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox meterrent_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox metertype_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   0
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Meter current reading"
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label connstate_lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Meter Already Connected"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter ID"
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Working"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Rent :"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Type :"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5535
      Left            =   3480
      Top             =   840
      Width           =   7935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   5775
      Left            =   3480
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "meter_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_meter As ADODB.Recordset
Public metertypeid As Long
Dim meterid As Long
Public searchedmeterID As Long
Public state As Integer

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
            qry = "delete from consumer_t where cid=" & cnsu_src_frm.cnsu_id & " "
            Call delete(qry)
                   
            MsgBox "Meter Record Deleted Successfully", vbInformation
          
            state = 3
            meterid_txt.Text = meterid
            metertype_txt.Text = ""
            meterrent_txt.Text = ""
            meterread_txt.Visible = False
            Label6.Visible = False
            meterread_txt = ""
            
            s_cmd.Enabled = False
            del_cmd.Enabled = False
            
            metertypsrc_cmd.Enabled = False
            yes_opt.Enabled = False
            no_opt.Enabled = False
            yes_opt.value = False
        End If
End If
End Sub

Private Sub Form_Load()
    Set rs_meter = New ADODB.Recordset
    rs_meter.CursorLocation = adUseClient
       
    rs_meter.Open ("select max(mid) from meter_t"), bms_cn, 3, 3
   ' rs_meter
    
    If rs_meter.RecordCount > 0 Then
         If IsNull(rs_meter.Fields(0)) Then
             meterid = 1
         Else
             meterid = rs_meter.Fields(0) + 1
         End If
    End If
       
    meterid_txt.Text = meterid
    rs_meter.Close
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub metertypsrc_cmd_Click()
metertype_src_frm.callvalue = 2
metertype_src_frm.Show vbModal

End Sub

Private Sub new_cmd_Click()
s_cmd.Enabled = True

metertypeid = -999
meterid_txt.Text = meterid
metertype_txt.Text = ""
meterrent_txt.Text = ""

metertypsrc_cmd.Enabled = True
yes_opt.Enabled = True
no_opt.Enabled = True

yes_opt.value = True

state = 1
End Sub

Private Sub no_opt_Click()
If no_opt.value = True Then
    meterread_txt.Visible = False
    Label6.Visible = False
    meterread_txt = ""
End If
End Sub

Private Sub s_cmd_Click()
    Select Case state
        Case 1
            If metertype_txt.Text <> "" Then
                If ((meterread_txt = "") And (yes_opt.value = True)) Then
                 MsgBox "Please Input meter reading", vbInformation
                 Exit Sub
                End If
                
                If (Len(meterread_txt) <> meterread_txt.MaxLength And (yes_opt.value = True)) Then
                 MsgBox "Please Input meter reading , this meter have " & meterread_txt.MaxLength & " digits", vbInformation
                 Exit Sub
                End If
                
                Dim qry As String
                Dim opt As Boolean
                If yes_opt.value = True Then
                    opt = 1
                    qry = "insert into meter_t values(" & meterid & ",'" & metertypeid & "','" & meterread_txt.Text & "',null," & opt & ")"
                ElseIf no_opt.value = True Then
                    opt = 0
                    qry = "insert into meter_t values(" & meterid & ",'" & metertypeid & "','',null," & opt & ")"
                End If
                
                
                Call insert(qry)
                MsgBox "New Record Saved Successfully", vbInformation
                state = 3
                
                metertypeid = -999
                meterid = meterid + 1
                meterid_txt.Text = meterid
                metertype_txt.Text = ""
                meterrent_txt.Text = ""
                meterread_txt = ""
                meterread_txt.Visible = False
                Label6.Visible = False
                meterread_txt = ""
                
                meterread_txt.Visible = False
                Label6.Visible = False
                metertypsrc_cmd.Enabled = False
                yes_opt.Enabled = False
                no_opt.Enabled = False
                yes_opt.value = False
                s_cmd.Enabled = False
            Else
                MsgBox "Please Search And Add Meter Type", vbInformation
            End If
        Case 2
                If yes_opt.value = True Then
                    opt = 1
                    qry = "update meter_t set  workstate=" & opt & ", mstartread = '" & meterread_txt.Text & "' where mid=" & searchedmeterID & " "
                ElseIf no_opt.value = True Then
                    opt = 0
                    qry = "update meter_t set workstate=" & opt & ", mstartread =NULL where mid=" & searchedmeterID & " "
                End If
                
                
                update (qry)
                state = 3
                
                meterid_txt.Text = meterid
                metertype_txt.Text = ""
                meterrent_txt.Text = ""
                meterread_txt = ""
            
                meterread_txt.Visible = False
                Label6.Visible = False
                
                s_cmd.Enabled = False
                del_cmd.Enabled = False
                
                metertypsrc_cmd.Enabled = False
                yes_opt.Enabled = False
                no_opt.Enabled = False
                yes_opt.value = False
                
                MsgBox "Record Update SuccessFully", vbInformation
        Case 3
    End Select
    
            
End Sub

Private Sub src_cmd_Click()
meter_src_frm.callvalue = 1
meter_src_frm.Show vbModal
End Sub

Private Sub yes_opt_Click()
If yes_opt.value = True Then
    meterread_txt = ""
    meterread_txt.Visible = True
    Label6.Visible = True
End If
End Sub
