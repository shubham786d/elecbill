VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tarif_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tariff Form"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "tarif_frm.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton fix_cmd 
      Caption         =   "Fix"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10920
      TabIndex        =   30
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox maxload_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8760
      MaxLength       =   255
      TabIndex        =   28
      Text            =   " "
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox minamt_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Text            =   " "
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox minload_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      MaxLength       =   255
      TabIndex        =   25
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   11160
      Picture         =   "tarif_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox mmc_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   18
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid tarifgrid 
      Height          =   2175
      Left            =   3840
      TabIndex        =   16
      Top             =   4320
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   8438015
      BackColorBkg    =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   8760
      Picture         =   "tarif_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   4320
      Picture         =   "tarif_frm.frx":15917
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      Picture         =   "tarif_frm.frx":15E44
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      Picture         =   "tarif_frm.frx":16411
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton rem_all_txt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      Picture         =   "tarif_frm.frx":16B25
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton rem_cmd 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Picture         =   "tarif_frm.frx":173C9
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton add_cmd 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Picture         =   "tarif_frm.frx":17B69
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox unitrate_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox range2_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox range1_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.ComboBox purpose_cmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox phase_cmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox ctype_cmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maximum Load"
      Height          =   375
      Left            =   7080
      TabIndex        =   29
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minimum Security Amount"
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minimum Load "
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit rate"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "range =>"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Monthly Unit:"
      Height          =   375
      Left            =   7080
      TabIndex        =   19
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<=range"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "connection type"
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purpose"
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phase given"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6735
      Left            =   2280
      Top             =   360
      Width           =   9975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   2280
      Top             =   240
      Width           =   9975
   End
End
Attribute VB_Name = "tarif_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_tar As ADODB.Recordset
Dim cmd_tar As ADODB.Command
Dim tarif_id As Long
Public fg_rowcount As Integer
Public state As Integer ' 1:insert 2:update
Dim fg_click_state As Boolean
Dim cmbvalue(3) As Integer
Public searched As Boolean ' false= not searched

Private Type fg_updatestate       ' for flexgrid update states in plc_ro_frm
     rangenum As Long
     oldrangenum As Long
     range1 As Long
     range2 As Long
     unitrate As Long
     datastate As Integer         '0=normal ; 1= inserted; 2=updated ; 3= delete
End Type


Private back_matrix() As fg_updatestate
Public backm_count As Integer


Public Function update_backmatrix(datastate As Long, range1 As Long, range2 As Long, rangenum As Long, unitrate As Long)
    Static var As Integer
    
    If datastate = -999 Then                     ' For Reseting all value of back grid
      ReDim fg_UState(0)
      var = 0
      backm_count = 0
      Exit Function
    End If
      
   ' MsgBox var
    ReDim Preserve back_matrix(var)
    back_matrix(var).datastate = datastate
    back_matrix(var).range1 = range1
    back_matrix(var).range2 = range2
    back_matrix(var).rangenum = rangenum
    back_matrix(var).oldrangenum = rangenum
    back_matrix(var).unitrate = unitrate
    Debug.Print back_matrix(var).datastate; back_matrix(var).range1; back_matrix(var).range2; back_matrix(var).rangenum; back_matrix(var).unitrate
     var = var + 1
End Function






Private Sub add_cmd_Click()
'On Error GoTo errorHandler
        If Val(range2_txt.Text) > Val(range1_txt.Text) Then
          If unitrate_txt <> "" Then
                 If tarifgrid.Rows = 1 Then       'fix setting
                    fix_cmd.Enabled = False
                    mmc_txt.Enabled = False
                 End If
                 
                 fg_rowcount = fg_rowcount + 1
                 tarifgrid.Rows = tarifgrid.Rows + 1
            
                 tarifgrid.TextMatrix(fg_rowcount, 0) = range1_txt.Text
                 tarifgrid.TextMatrix(fg_rowcount, 1) = range2_txt.Text
                 tarifgrid.TextMatrix(fg_rowcount, 2) = unitrate_txt.Text
                 tarifgrid.TextMatrix(fg_rowcount, 3) = fg_rowcount
                 ReDim Preserve back_matrix(backm_count)
                 
                 'If state = 1 Then
                     back_matrix(backm_count).range1 = range1_txt.Text
                     back_matrix(backm_count).range2 = range2_txt.Text
                     back_matrix(backm_count).unitrate = unitrate_txt.Text
                     back_matrix(backm_count).rangenum = fg_rowcount
                     back_matrix(backm_count).datastate = 1
                     backm_count = backm_count + 1
               '  End If
                 
                 range1_txt.Text = range2_txt.Text + 1
                 range2_txt.Text = ""
                 unitrate_txt.Text = ""
            
    '                    ReDim Preserve fg_roundsid(fg_rowcount)
    '                    fg_roundsid(fg_rowcount) = ro_cmb.ItemData(ro_cmb.ListIndex)
    '                    rounds_fg.Rows = rounds_fg.Rows + 1
    '                    fg_rowcount = fg_rowcount + 1
            Else
                 MsgBox "Please Input Unit Rate", vbInformation
            End If
        Else
            MsgBox "Range 2 can note be Smaller than range 1", vbInformation
        End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub co_ex_cmd_Click()
        If bms_mdi.tabclose_flag = False Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub Command1_Click()
Dim i As Integer
        For i = 0 To UBound(back_matrix)
        Form1.MSFlexGrid1.TextMatrix(i, 0) = back_matrix(i).rangenum
        Form1.MSFlexGrid1.TextMatrix(i, 1) = back_matrix(i).oldrangenum
        Form1.MSFlexGrid1.TextMatrix(i, 2) = back_matrix(i).range1
        Form1.MSFlexGrid1.TextMatrix(i, 3) = back_matrix(i).range2
        Form1.MSFlexGrid1.TextMatrix(i, 4) = back_matrix(i).unitrate
        Form1.MSFlexGrid1.TextMatrix(i, 5) = back_matrix(i).datastate
        
        Form1.MSFlexGrid1.Rows = Form1.MSFlexGrid1.Rows + 1
        Next
        Form1.Show
End Sub



Private Sub ctype_cmb_Click()
Call check_all
End Sub

Private Sub del_cmd_Click()
If state = 2 Then
        Dim test As Integer
        test = MsgBox("Do U Want To Delete This Record ?", vbYesNoCancel + vbQuestion, "Information")
         
         If test = 6 Then
            Dim qry As String
            qry = "delete from tarif_t where tarifid=" & tarif_src_frm.tarifid & "  "
            Call delete(qry)
            qry = "delete from tarifsetting_t where tarifid=" & tarif_src_frm.tarifid & "  "
            Call delete(qry)
            
            MsgBox "tarif Record Deleted Successfully", vbInformation
    
            state = 3
            Call formreset
        End If
End If
End Sub

Private Sub fix_cmd_Click()
    If mmc_txt.Text <> "" Then
        range1_txt.Text = mmc_txt.Text
        range2_txt.Enabled = True
        unitrate_txt.Enabled = True
        add_cmd.Enabled = True
        rem_cmd.Enabled = True
        rem_all_txt.Enabled = True
    Else
        MsgBox "Please Input minimum monthly unit", vbInformation
    End If
End Sub

Private Sub Form_Load()
    range1_txt.Text = "0"
    range2_txt.Text = ""
    unitrate_txt.Text = ""

    ReDim back_matrix(0)
       fg_rowcount = 0
       backm_count = 0
       tarifgrid.FixedCols = 0
       
       If tarifgrid.Rows < 1 Then
       tarifgrid.FixedRows = 1
       End If
       
       tarifgrid.Cols = 4
       tarifgrid.Rows = 1
       
       tarifgrid.TextMatrix(fg_rowcount, 0) = "Range1"
       tarifgrid.TextMatrix(fg_rowcount, 1) = "Range2 "
       tarifgrid.TextMatrix(fg_rowcount, 2) = "Rate"
       tarifgrid.TextMatrix(fg_rowcount, 3) = "S.no"
       

    Set rs_tar = New ADODB.Recordset
    rs_tar.CursorLocation = adUseClient
       
    rs_tar.Open ("select max(tarifid) from tarif_t"), bms_cn, 3, 3
   ' rs_con
    
    If rs_tar.RecordCount > 0 Then
         If IsNull(rs_tar.Fields(0)) Then
             tarif_id = 1
         Else
             tarif_id = rs_tar.Fields(0) + 1
         End If
    End If
       
    'MsgBox tarif_id
    rs_tar.Close


    Dim query As String
    
    query = "select * from perpose_t"
    
    Call setcombo(query, purpose_cmb, "--Select Purpose--", 1, 0)
    
    query = "select * from phase_t"
    
    Call setcombo(query, phase_cmb, "--Select Phase--", 1, 0)
    
    query = "select * from type_t"
    
    Call setcombo(query, ctype_cmb, "--Select Type--", 1, 0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
         If (bms_mdi.tabclose_flag = False) And (bms_mdi.alltabexit = False) Then
            Call bms_mdi.tab_close
        End If
    
        If bms_mdi.t_count = 0 Then
            bms_mdi.Picture4.Height = 0
        End If
End Sub

Private Sub maxload_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub minamt_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub minload_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub mmc_txt_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub

Private Sub new_cmd_Click()
    Call Form_Load
    state = 1
    range1_txt.Text = "0"
    range2_txt.Text = ""
    unitrate_txt.Text = ""
    minload_txt.Text = ""
    maxload_txt.Text = ""
    minamt_txt.Text = ""
    
    tarifgrid.FixedCols = 0
    ReDim back_matrix(0)
    backm_count = 0
    fg_click_state = False
    fg_rowcount = 0
    
    
    phase_cmb.ListIndex = 0
    purpose_cmb.ListIndex = 0
    ctype_cmb.ListIndex = 0
    phase_cmb.Enabled = True
    purpose_cmb.Enabled = True
    ctype_cmb.Enabled = True
    minload_txt.Enabled = True
    maxload_txt.Enabled = True
    minamt_txt.Enabled = True
    mmc_txt.Enabled = True
    s_cmd.Enabled = True
    
    Command1.Enabled = True
    tarifgrid.Enabled = True
    fix_cmd.Enabled = True
    
    'Cancel_cmd.Enabled = True
    
    
    'co_cmb.Enabled = False
End Sub



Private Sub phase_cmb_Click()
Call check_all
End Sub





Private Sub purpose_cmb_Click()
Call check_all
End Sub

Private Sub rem_cmd_Click()
    If tarifgrid.Rows > 1 Then
        If fg_click_state = True Then
            Dim i, k As Integer
            For k = 0 To UBound(back_matrix)            ' for back remove
                If tarifgrid.TextMatrix(tarifgrid.Row, 0) = back_matrix(k).range1 And tarifgrid.TextMatrix(tarifgrid.Row, 1) = back_matrix(k).range2 And back_matrix(k).datastate <> 3 Then
                    If (k <> UBound(back_matrix)) Then
                      back_matrix(k + 1).range1 = back_matrix(k).range1
                    End If
                    
                    back_matrix(k).datastate = 3           ' removed marked
                    Exit For
                End If
            Next
            
            If tarifgrid.Row = 1 Then
                tarifgrid.TextMatrix(tarifgrid.Row, 0) = Val(mmc_txt.Text)
            End If
            
            
            
            For i = tarifgrid.Row To tarifgrid.Rows - 2    ' for back display
              If (tarifgrid.Row <> 1) Then
                tarifgrid.TextMatrix(i, 0) = tarifgrid.TextMatrix(i - 1, 1) + 1
              End If
                tarifgrid.TextMatrix(i, 1) = tarifgrid.TextMatrix(i + 1, 1)
                tarifgrid.TextMatrix(i, 2) = tarifgrid.TextMatrix(i + 1, 2)
            Next
            
            If tarifgrid.Rows - 1 = 0 Then                  'for front reset range
                range1_txt.Text = Val(mmc_txt.Text)
            ElseIf tarifgrid.Row = tarifgrid.Rows - 1 Then
                range1_txt.Text = tarifgrid.TextMatrix(tarifgrid.Rows - 1, 0)
            End If

            
            Dim numrange As Integer
            numrange = 1
            
            If state = 2 Then                                  ' // for back range num set //
                For k = 0 To UBound(back_matrix)
                    If back_matrix(k).datastate <> 3 Then
                        back_matrix(k).rangenum = numrange
                        
                        If back_matrix(k).datastate <> 1 Then  'not for inserted
                            back_matrix(k).datastate = 2
                        End If
                        numrange = numrange + 1
                    End If
                Next
            End If
            
            If tarifgrid.Rows = 2 Then     'fix setting
                    fix_cmd.Enabled = True
                    mmc_txt.Enabled = True
                    
                    range1_txt.Text = mmc_txt.Text
                    
                    range2_txt.Enabled = False
                    unitrate_txt.Enabled = False
                    add_cmd.Enabled = False
                    rem_cmd.Enabled = False
                    rem_all_txt.Enabled = False
            End If
                 
            tarifgrid.Rows = tarifgrid.Rows - 1
            fg_rowcount = fg_rowcount - 1
         Else
             MsgBox "Please Select Any Record From Table For Removing", vbInformation
        End If
    End If
    
    fg_click_state = False
End Sub

Private Sub s_cmd_Click()
     
    If isduplicate() = False Then
        If phase_cmb.ListIndex <> 0 Then
            If purpose_cmb.ListIndex <> 0 Then
                If ctype_cmb.ListIndex <> 0 Then
                    If minload_txt <> "" Then
                        If maxload_txt.Text <> "" Then
                            If Val(minload_txt.Text) < Val(maxload_txt.Text) Then
                                If minamt_txt <> "" Then
                                    If fg_rowcount <> 0 Then
                                        Set cmd_tar = New ADODB.Command
                                        cmd_tar.CommandType = adCmdText
                                        Dim i As Integer
                                        
                                        If state = 1 And tarifgrid.Rows > 1 Then
                                           Dim qry As String
                                           qry = "insert into tarif_t values(" & tarif_id & ",'" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & "','" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & "','" & phase_cmb.ItemData(phase_cmb.ListIndex) & "','" & mmc_txt.Text & "','" & minload_txt.Text & "','" & maxload_txt.Text & "','" & minamt_txt.Text & "')"
                                           Call insert(qry)
                                        ElseIf state = 2 Then
                                           qry = "update tarif_t set typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & ",perposid=" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & ",phaseid=" & phase_cmb.ItemData(phase_cmb.ListIndex) & ",mmcprice=" & mmc_txt.Text & " where tarifid=" & tarif_src_frm.tarifid & ""
                                           Call update(qry)
                                        End If
                                        
                                        
                                        For i = 0 To UBound(back_matrix)
                                           Select Case back_matrix(i).datastate
                                             Case 0
                                             Case 3  'deletion
                                                If state = 2 Then
                                                
                                                qry = " delete * from tarifsetting_t where tarifid=" & tarif_src_frm.tarifid & " and rangenum=" & back_matrix(i).oldrangenum & ""
                                                Call delete(qry)
                                                End If
                                             Case 1  'insertion
                                                Dim str As String
                                                If state = 1 Then
                                                str = "insert into tarifsetting_t values(" & tarif_id & ",'" & back_matrix(i).rangenum & "','" & back_matrix(i).range1 & "','" & back_matrix(i).range2 & "','" & back_matrix(i).unitrate & "')"
                                                Else
                                                str = "insert into tarifsetting_t values(" & tarif_src_frm.tarifid & ",'" & back_matrix(i).rangenum & "','" & back_matrix(i).range1 & "','" & back_matrix(i).range2 & "','" & back_matrix(i).unitrate & "')"
                                                End If
                                        
                                                Call insert(str)
                                            Case 2 'update
                                                str = "update tarifsetting_t set range1=" & back_matrix(i).range1 & ",range2=" & back_matrix(i).range2 & " , rangenum=" & back_matrix(i).rangenum & " where tarifid=" & tarif_src_frm.tarifid & " and rangenum=" & back_matrix(i).oldrangenum & " "
                                                Call update(str)
                                           End Select
                                        Next
                                           
                                           
                                           If state = 1 Then
                                            tarif_id = tarif_id + 1
                                            MsgBox " Records Saved Successfully", vbInformation
                                           Else
                                            MsgBox " Records Updated Successfully", vbInformation
                                           End If
                                           Call formreset
                                    Else
                                        MsgBox "Please Input some Tarif's in grid ", vbInformation
                                    End If
                                Else
                                    MsgBox "Please input minimum security amount ", vbInformation
                                End If
                            Else
                                MsgBox "Maximum load can not be less than minimum load ", vbInformation
                            End If
                        Else
                            MsgBox "please input Maximum  load value ", vbInformation
                        End If
                    Else
                        MsgBox "Please input minimum  load value  ", vbInformation
                    End If
                Else
                    MsgBox "Please select connection type  ", vbInformation
                End If
            Else
                  MsgBox "Please select purpose ", vbInformation
            End If
        Else
            MsgBox "Please select tariff Phase", vbInformation
        End If
    Else
        MsgBox "tarif settings of selected type is alreay existing. Please select other types"
    End If
End Sub

Private Sub src_cmd_Click()
tarif_src_frm.callvalue = 1
tarif_src_frm.Show vbModal
End Sub

Private Sub tarifgrid_Click()
If tarifgrid.Row > 0 Then
        fg_click_state = True
        'fg_updaterow = elig_fg.Row
    End If
End Sub

Private Sub formreset()
        range1_txt.Text = "0"
        range2_txt.Text = ""
        unitrate_txt.Text = ""

    
        tarifgrid.FixedCols = 0
        
        
        ReDim back_matrix(0)
        backm_count = 0
        fg_click_state = False
        state = 3
        fg_rowcount = 0
        
        phase_cmb.ListIndex = 0
        purpose_cmb.ListIndex = 0
        ctype_cmb.ListIndex = 0
        
        phase_cmb.Enabled = False
        purpose_cmb.Enabled = False
        ctype_cmb.Enabled = False
        range2_txt.Enabled = False
        unitrate_txt.Enabled = False
        mmc_txt.Enabled = False
        s_cmd.Enabled = False
        
        Command1.Enabled = False
        tarifgrid.Enabled = False
        
        add_cmd.Enabled = False
        rem_cmd.Enabled = False
        rem_all_txt.Enabled = False
        del_cmd.Enabled = False
        
        minload_txt.Text = ""
        maxload_txt.Text = ""
        minamt_txt.Text = ""
        mmc_txt.Text = ""
        
        minload_txt.Enabled = False
        maxload_txt.Enabled = False
        minamt_txt.Enabled = False
        tarifgrid.Rows = 1
End Sub

Private Function isduplicate() As Boolean
        Set rs_tar = New ADODB.Recordset
        rs_tar.CursorLocation = adUseClient
        
        rs_tar.Open ("select * from tarif_t where typeid=" & ctype_cmb.ItemData(ctype_cmb.ListIndex) & " and perposid=" & purpose_cmb.ItemData(purpose_cmb.ListIndex) & " and phaseid=" & phase_cmb.ItemData(phase_cmb.ListIndex) & ""), bms_cn, 3, 3
        
        If state = 1 Then
            If rs_tar.RecordCount >= 1 Then
              isduplicate = True
            ElseIf rs_tar.RecordCount = 0 Then
              isduplicate = False
            End If
        Else
            If cmbvalue(0) = ctype_cmb.ItemData(ctype_cmb.ListIndex) And purpose_cmb.ItemData(purpose_cmb.ListIndex) = cmbvalue(2) And phase_cmb.ItemData(phase_cmb.ListIndex) = cmbvalue(1) Then
                isduplicate = False
            ElseIf rs_tar.RecordCount >= 1 Then
                isduplicate = True
            End If
        End If
            
        'Debug.Print rs_tar
        
End Function

Public Sub sethistoryvalueofcmb(v1 As Integer, v2 As Integer, v3 As Integer)
cmbvalue(0) = v1
cmbvalue(1) = v2
cmbvalue(2) = v3
End Sub


Private Sub check_all() ' for combo validation
    If ((ctype_cmb.ListIndex <> 0) And (purpose_cmb.ListIndex <> 0) And (phase_cmb.ListIndex <> 0)) Then
        If isduplicate() = True And searched = False Then
            MsgBox "tarif settings of selected type is alreay existing. Please select other types"
        Else
            searched = False
        End If
    End If
End Sub

