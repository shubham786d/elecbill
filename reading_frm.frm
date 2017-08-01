VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form reading_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reading frm"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "reading_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ivrssrc_cmd 
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   255
   End
   Begin MSComCtl2.DTPicker tempdate 
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Format          =   11993089
      CurrentDate     =   42537
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   8520
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker readtaken_dtp 
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   11993091
      CurrentDate     =   42507
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      Picture         =   "reading_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      Picture         =   "reading_frm.frx":154CB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   4200
      Picture         =   "reading_frm.frx":15A98
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   8640
      Picture         =   "reading_frm.frx":15FC5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox creading_txt 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox name_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox ivrs_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      MaxLength       =   255
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   11400
      Picture         =   "reading_frm.frx":164B3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid readinggrid 
      Height          =   3615
      Left            =   2520
      TabIndex        =   7
      Top             =   2640
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   8438015
      BackColorBkg    =   16777215
      Enabled         =   0   'False
      AllowUserResizing=   1
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
   Begin MSComCtl2.DTPicker readingofmonth_dtp 
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "MMMM/yyyy"
      Format          =   11993091
      CurrentDate     =   42507
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reading of month"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of Reading taken"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IVRS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6855
      Left            =   2280
      Top             =   240
      Width           =   9975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   7095
      Left            =   2280
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label Label2 
      Caption         =   "Minimum Unit"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8640
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "reading_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public state As Integer
Public ivrs As Long
Public fg_rowcount As Integer
Dim fg_click_state As Boolean
Dim reading_id As Long

Dim rs_read As ADODB.Recordset
Dim rsr As ADODB.Recordset
Dim Recordsett As Recordset

Dim minmonth As String 'min read month
Dim minunit As Long
Dim readingdatedflag As Integer '0 = normal ,, 1= less read date ,, 2=more read date


Const Checked = "þ"
Const UnChecked = "q"



Private Sub Check1_Click()
    If preading_txt <> "" Then
        If (Check1.value = 1) Then
            creading_txt.Enabled = False
            creading_txt.Text = minunit + preading_txt
            
            Do While Len(creading_txt.Text) <> creading_txt.MaxLength
                creading_txt.Text = "0" + creading_txt.Text
            Loop
        Else
             creading_txt.Enabled = True
             creading_txt.Text = ""
        End If
    End If
End Sub



Private Sub Command1_Click()

End Sub

Private Sub creading_txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       
            
            While Len(creading_txt) <> creading_txt.MaxLength
                creading_txt = "0" + creading_txt
            Wend
            
            If Val(readinggrid.TextMatrix(readinggrid.Row, 1)) <= Val(creading_txt) Then
                Set rs_read = New ADODB.Recordset
                rs_read.CursorLocation = adUseClient
                Debug.Print "select * from reading_t where readingofmonth like ' " & DateAdd("m", 1, readingofmonth_dtp) & "' "
                rs_read.Open "select * from reading_t where readingofmonth like '" & DateAdd("m", 1, readingofmonth_dtp) & "' and ivrs=" & readinggrid.TextMatrix(readinggrid.Row, 0) & " ", bms_cn, 3, 3

                '/// next month current can not be grater
                If rs_read.RecordCount <> 0 Then
                    If creading_txt > rs_read.Fields(5) Then
                        MsgBox "Current Reading of This Month can not be More then Next Month Reading which is " & rs_read.Fields(5) & "", vbInformation
                        Exit Sub
                    End If
                End If
                
                If readinggrid.TextMatrix(readinggrid.Row, 3) = Checked Then
                    readinggrid.TextMatrix(readinggrid.Row, 3) = UnChecked
                End If
                
                
                '//// minimum (anklat kapt)
                If Val(creading_txt) < (Val(readinggrid.TextMatrix(readinggrid.Row, 1)) + Val(readinggrid.TextMatrix(readinggrid.Row, 4))) Then
                    readinggrid.TextMatrix(readinggrid.Row, 3) = Checked
                    readinggrid.TextMatrix(readinggrid.Row, 2) = Val(readinggrid.TextMatrix(readinggrid.Row, 1)) + Val(readinggrid.TextMatrix(readinggrid.Row, 4))
                    
                    Do While Len(readinggrid.TextMatrix(readinggrid.Row, 2)) < Len(readinggrid.TextMatrix(readinggrid.Row, 1))
                         readinggrid.TextMatrix(readinggrid.Row, 2) = "0" + readinggrid.TextMatrix(readinggrid.Row, 2)
                    Loop
                     creading_txt.Text = ""
                     creading_txt.Visible = False
                Else
                    readinggrid.TextMatrix(readinggrid.Row, 2) = creading_txt
                    creading_txt.Visible = False
                    creading_txt = ""
                End If
            Else
                MsgBox "Current reading can not Be less Than Previous Reading"
            End If
        
    Else
        Select Case KeyAscii
                Case 48 To 57 'numaric
                Case 8      'backspace
                Case Else
                  KeyAscii = 0
        End Select
    End If
    
End Sub

Private Sub creading_txt_LostFocus()
creading_txt.Visible = False
creading_txt = ""
End Sub

Private Sub Form_Load()
       readinggrid.Cols = 7
       readinggrid.Rows = 1
       readinggrid.ColWidth(0) = 2000
       readinggrid.ColWidth(1) = 2000
       readinggrid.ColWidth(2) = 2000
       readinggrid.ColWidth(3) = 2000
       
       readinggrid.FixedCols = 0
       
       fg_rowcount = 0
       
       readinggrid.TextMatrix(fg_rowcount, 0) = "Ivrs"
       readinggrid.TextMatrix(fg_rowcount, 1) = "Previous Reading"
       readinggrid.TextMatrix(fg_rowcount, 2) = "Current Reading"
       readinggrid.TextMatrix(fg_rowcount, 3) = "Minimum Reading"
       readinggrid.TextMatrix(fg_rowcount, 4) = "Billed"
       readinggrid.TextMatrix(fg_rowcount, 5) = "reading ID"
    Set rs_read = New ADODB.Recordset
    rs_read.CursorLocation = adUseClient
       
    rs_read.Open ("select max(rid) from reading_t"), bms_cn, 3, 3
   
    
    If rs_read.RecordCount > 0 Then
         If IsNull(rs_read.Fields(0)) Then
             reading_id = 1
         Else
             reading_id = rs_read.Fields(0) + 1
         End If
    End If
    
    Set rs_read = New ADODB.Recordset
    rs_read.CursorLocation = adUseClient
       
    rs_read.Open ("select * from settings_t"), bms_cn, 3, 3
    
     readingofmonth_dtp.value = rs_read.Fields(0)
     minmonth = readingofmonth_dtp.value
     
    Set rs_read = New ADODB.Recordset
    rs_read.Open "select Max(readingofmonth) from reading_t", bms_cn, 3, 3
    
    If (IsNull(rs_read.Fields(0)) = False) Then
        tempdate.value = CDate(rs_read.Fields(0))
        tempdate.value = DateAdd("m", 1, tempdate)
        If tempdate.month <> 12 Then
            tempdate.month = tempdate.month Mod 12
        End If
        
        readingofmonth_dtp.value = tempdate.value
    End If
     
    readtaken_dtp.CustomFormat = "dd/MM/yyyy"
    readtaken_dtp.Format = dtpCustom
End Sub




Private Sub ivrssrc_cmd_Click()
con_src_frm.cldform = 3
con_src_frm.Show vbModal

End Sub





Private Sub ivrs_txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        For i = 1 To readinggrid.Rows - 1
            If readinggrid.TextMatrix(i, 0) = ivrs_txt.Text Then
                readinggrid.Row = i
                readinggrid.RowSel = i
                readinggrid.Col = 0
                readinggrid.ColSel = readinggrid.Cols - 1
                Exit Sub
            End If
        Next
        MsgBox "IVRS Not Found In List"
    End If

End Sub


Private Sub new_cmd_Click()
state = 1
readtaken_dtp.Enabled = True
readingofmonth_dtp.Enabled = True
s_cmd.Enabled = True

creading_txt.Enabled = True

'add_cmd.Enabled = True
'rem_cmd.Enabled = True
ivrs_txt.Enabled = True
'ivrssrc_cmd.Enabled = True
readinggrid.Enabled = True
End Sub

Private Sub preading_txt_Change()

End Sub

Private Sub readinggrid_Click()

    If readinggrid.Row > 0 Then
        fg_click_state = True
        
        If (readinggrid.Col = 3) Then   'for minimum chked and uncked
            If readinggrid.TextMatrix(readinggrid.Row, readinggrid.Col) = UnChecked Then
                readinggrid.TextMatrix(readinggrid.Row, readinggrid.Col) = Checked
                readinggrid.TextMatrix(readinggrid.Row, 2) = Val(readinggrid.TextMatrix(readinggrid.Row, 1)) + Val(readinggrid.TextMatrix(readinggrid.Row, 4))
               Do While Len(readinggrid.TextMatrix(readinggrid.Row, 2)) < Len(readinggrid.TextMatrix(readinggrid.Row, 1))
                    readinggrid.TextMatrix(readinggrid.Row, 2) = "0" + readinggrid.TextMatrix(readinggrid.Row, 2)
               Loop
                creading_txt.Text = ""
                creading_txt.Visible = False
            Else
                readinggrid.TextMatrix(readinggrid.Row, readinggrid.Col) = UnChecked
                readinggrid.TextMatrix(readinggrid.Row, 2) = ""
                creading_txt.Text = ""
                creading_txt.Visible = False
            End If
        End If
        
        If (readinggrid.Col = 2) Then   ' for move of textbox
            creading_txt.Move readinggrid.Left + readinggrid.CellLeft, readinggrid.Top + readinggrid.CellTop
            creading_txt.MaxLength = 255
            creading_txt = readinggrid.TextMatrix(readinggrid.Row, 2)
            creading_txt.MaxLength = Len(readinggrid.TextMatrix(readinggrid.Row, 1))
            creading_txt.Width = readinggrid.CellWidth
            creading_txt.Visible = True
            creading_txt.SetFocus
        End If
    End If
End Sub

Private Sub readingofmonth_dtp_CloseUp()
    
    
    Dim actdate As String
    readingofmonth_dtp.Day = 1

    actdate = readingofmonth_dtp.value

    Set rs_read = New ADODB.Recordset
    Dim readmonth As String



'// fixing month view
'    If readingofmonth_dtp.month < 10 Then
'       readmonth = "0" & readingofmonth_dtp.month
'    Else
       readmonth = readingofmonth_dtp.month
   ' End If


'// search for reading on selected month

    rs_read.CursorLocation = adUseClient
    Debug.Print "select * from reading_t where DatePart('m', [readingofmonth]) = " & readmonth & ""
    rs_read.Open "select * from reading_t where DatePart('m', [readingofmonth]) = " & readmonth & " and DatePart('yyyy', [readingofmonth]) = " & readingofmonth_dtp.year & "", bms_cn, 3, 3

'// if already taken
    If (rs_read.RecordCount <> 0) Then
        MsgBox "Selected Month reading Is Already Taken"
        Set rs_read = New ADODB.Recordset
        rs_read.Open "select Max(readingofmonth) from reading_t", bms_cn, 3, 3
        
        If (Not IsNull(rs_read.Fields(0))) Then
            tempdate.value = CDate(rs_read.Fields(0))
            tempdate.value = DateAdd("m", 1, tempdate)
            readingofmonth_dtp.value = tempdate.value
        End If
'// else if less then admin date
    ElseIf readingofmonth_dtp.value < CDate(minmonth) Then
        MsgBox "Reading can not be taken from less than starting of  reading date set by ADMIN"
        readingofmonth_dtp.value = minmonth
        readingofmonth_dtp.Day = 1
'// else
    Else
        Set rs_read = New ADODB.Recordset
        rs_read.Open "select Max(readingofmonth) from reading_t", bms_cn, 3, 3
        
        If (Not IsNull(rs_read.Fields(0))) Then
                tempdate.value = CDate(rs_read.Fields(0))
                tempdate.value = DateAdd("m", 1, tempdate)

                If readingofmonth_dtp.value <> tempdate.value Then
                    MsgBox " Previous month reading is not taken "
                    readingofmonth_dtp.value = tempdate.value
                    Exit Sub
                End If
         Else
            Set rs_read = New ADODB.Recordset
            rs_read.Open "select * from settings_t", bms_cn, 3, 3
            
            If CDate(rs_read.Fields(0)) <> CDate(readingofmonth_dtp.value) Then
                MsgBox " Previous month reading is not taken "
                Exit Sub
            End If
         End If
        
            
            Set rs_read = New ADODB.Recordset 'geting ivrs
            Debug.Print "select * from connection_t where cdate <= 1/" & readmonth & "/" & readingofmonth_dtp.year & ""
            rs_read.Open "select * from connection_t where cdate <= #" & readmonth & "/15/" & readingofmonth_dtp.year & "#", bms_cn, 3, 3
            
            Dim i As Long
            fg_rowcount = 0
            readinggrid.Rows = 1
            
            For i = 1 To rs_read.RecordCount
                fg_rowcount = fg_rowcount + 1
                readinggrid.Rows = readinggrid.Rows + 1
                readinggrid.TextMatrix(i, 0) = rs_read.Fields(1)
                 Set rsr = New ADODB.Recordset   'geting Pre reading
                 rsr.Open "select * from meter_t where mid =" & rs_read.Fields(2) & "  ", bms_cn, 3, 3
                 ivrs_txt.Enabled = True
                readinggrid.TextMatrix(i, 1) = rsr.Fields(2)
                
                readinggrid.Row = i
                readinggrid.Col = 3
                readinggrid.CellFontName = "Wingdings"
                readinggrid.CellFontSize = 14
                readinggrid.CellAlignment = flexAlignCenterCenter
                readinggrid.Text = UnChecked

                Set Recordsett = New ADODB.Recordset ' geting min unit
                Recordsett.Open "select * from tarif_t where tarifid =" & rs_read.Fields(3) & "  ", bms_cn, 3, 3
                 
                readinggrid.TextMatrix(i, 4) = Recordsett.Fields(4)
                
                rs_read.MoveNext
            Next
    End If
End Sub

Private Sub readtaken_dtp_CloseUp()

        If readtaken_dtp.value < readingofmonth_dtp.value Then
            MsgBox "Reading Date can not be less than Reading of Month"
            
            If (month(readingofmonth_dtp.value) + 1) > 12 Then
                readtaken_dtp.month = (month(readingofmonth_dtp.value) + 1) Mod 12
                readtaken_dtp.year = readtaken_dtp.year + 1
            Else
                readtaken_dtp.month = month(readingofmonth_dtp.value) + 1
            End If
        ElseIf (readtaken_dtp.month = readingofmonth_dtp.month) And (readtaken_dtp.year = readingofmonth_dtp.year) Then
            If totaldayinmonth(readtaken_dtp.month, readtaken_dtp.year) <> readtaken_dtp.Day Then
                MsgBox "Reading Date can not be equal to Reading of Month"
            End If
        ElseIf readtaken_dtp.value > readingofmonth_dtp.value Then
            
            Dim totalday As Integer
            
            If (month(readingofmonth_dtp.value) + 1) > 12 Then
                totalday = totaldayinmonth((month(readingofmonth_dtp.value) + 1) Mod 12, readtaken_dtp.year + 1)
                If CDate("" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) Mod 12 & "/" & readtaken_dtp.year + 1 & "") < readtaken_dtp.value Then
                    MsgBox "Reading Date can Not be More than 1 Month"
                    readtaken_dtp.Day = 1
                    readtaken_dtp.month = (month(readingofmonth_dtp.value) + 1)
                End If
            Else
                totalday = totaldayinmonth(month(readingofmonth_dtp.value) + 1, readingofmonth_dtp.year)
                Debug.Print "" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) & "/" & readtaken_dtp.year & ""
                If CDate("" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) & "/" & readtaken_dtp.year & "") < CDate(readtaken_dtp.value) Then
                    MsgBox "Reading Date can Not be More than 1 Month"
                    readtaken_dtp.Day = 1
                    readtaken_dtp.month = (month(readingofmonth_dtp.value) + 1)
                End If
            End If
        End If
    
End Sub

Private Function totaldayinmonth(mn As Integer, year As Integer) As Integer
    mn = mn Mod 12
    If mn = 0 Then
        mn = 1
    End If
    
    Select Case mn
        Case 1, 3, 5, 7, 8, 10, 12
            totaldayinmonth = 31
        Case 4, 6, 9, 11
            totaldayinmonth = 30
        Case 2
            If (((year Mod 100) = 0) And ((year Mod 400) = 0)) Or ((year Mod 4) = 0) Then
              totaldayinmonth = 29
            Else
              totaldayinmonth = 28
            End If
    End Select
    
End Function


Private Sub rem_cmd_Click()
    If fg_click_state = True Then
            Dim i, k As Integer
            For k = 0 To UBound(back_matrix)            ' for backgrid remove
                If readinggrid.TextMatrix(readinggrid.Row, 0) = back_matrix(k).ivrs And back_matrix(k).datastate <> 3 Then
                    back_matrix(k).datastate = 3           ' removed marked
                    Exit For
                End If
            Next
            
            
            For i = readinggrid.Row To readinggrid.Rows - 2    ' for back display
                readinggrid.TextMatrix(i, 0) = readinggrid.TextMatrix(i + 1, 0)
                readinggrid.TextMatrix(i, 1) = readinggrid.TextMatrix(i + 1, 1)
            Next
            readinggrid.Rows = readinggrid.Rows - 1
    End If
End Sub

Private Sub s_cmd_Click()
    Select Case state
        Case 1
              
                If readtaken_dtp.value < readingofmonth_dtp.value Then
                    MsgBox "Reading Date can not be less than Reading of Month"
                    Exit Sub
                ElseIf (readtaken_dtp.month = readingofmonth_dtp.month) And (readtaken_dtp.year = readingofmonth_dtp.year) Then
                    If totaldayinmonth(readtaken_dtp.month, readtaken_dtp.year) <> readtaken_dtp.Day Then
                        MsgBox "Reading Date can not be equal to Reading of Month"
                        Exit Sub
                    End If
                ElseIf readtaken_dtp.value > readingofmonth_dtp.value Then
            
                    Dim totalday As Integer
            
                    If (month(readingofmonth_dtp.value) + 1) > 12 Then
                        totalday = totaldayinmonth((month(readingofmonth_dtp.value) + 1) Mod 12, readtaken_dtp.year + 1)
                        If CDate("" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) Mod 12 & "/" & readtaken_dtp.year + 1 & "") < readtaken_dtp.value Then
                            MsgBox "Reading Date can Not be More than 1 Month"
                            readtaken_dtp.Day = 1
                            Exit Sub
                        End If
                    Else
                        totalday = totaldayinmonth(month(readingofmonth_dtp.value) + 1, readingofmonth_dtp.year)
                        Debug.Print "" & totalday & " / " & (month(readingofmonth_dtp.value) + 1) & " / " & readtaken_dtp.year & ""
                        If CDate("" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) & "/" & readtaken_dtp.year & "") < readtaken_dtp.value Then
                            MsgBox "Reading Date can Not be More than 1 Month"
                            readtaken_dtp.Day = 1
                            Exit Sub
                        End If
                    End If
                End If
                
        
              Dim i As Integer
              Dim qry As String
              Dim min As Integer '0 = unchek; 1=check
              
     
                For i = 1 To readinggrid.Rows - 1
                    If readinggrid.TextMatrix(i, 2) = "" Then
                        MsgBox "Please Fill All Readings"
                        Exit Sub
                    End If
                Next
              
              
              For i = 1 To readinggrid.Rows - 1
                If readinggrid.TextMatrix(i, 3) = UnChecked Then
                  min = 0
                Else
                  min = 1
                End If
        
                qry = "insert into reading_t values(" & reading_id & "," & Val(readinggrid.TextMatrix(i, 0)) & ", '" & str(readtaken_dtp.value) & "','" & str(readingofmonth_dtp.value) & "' ,'" & readinggrid.TextMatrix(i, 1) & "','" & readinggrid.TextMatrix(i, 2) & "',0," & min & ")"
                Call insert(qry)
                 qry = "update meter_t set mstartread='" & readinggrid.TextMatrix(i, 2) & "' where mid=(select meter_id from connection_t where ivrs='" & Val(readinggrid.TextMatrix(i, 0)) & "' ) "
                Call insert(qry)
                reading_id = reading_id + 1
              Next
              MsgBox "New Record Saved Successfully", vbInformation
              Call form_reset
        Case 2
               If readtaken_dtp.value < readingofmonth_dtp.value Then
                    MsgBox "Reading Date can not be less than Reading of Month"
                    Exit Sub
                ElseIf (readtaken_dtp.month = readingofmonth_dtp.month) And (readtaken_dtp.year = readingofmonth_dtp.year) Then
                    If totaldayinmonth(readtaken_dtp.month, readtaken_dtp.year) <> readtaken_dtp.Day Then
                        MsgBox "Reading Date can not be equal to Reading of Month"
                        Exit Sub
                    End If
                ElseIf readtaken_dtp.value > readingofmonth_dtp.value Then
            
                   '  Dim totalday As Integer
            
                    If (month(readingofmonth_dtp.value) + 1) > 12 Then
                        totalday = totaldayinmonth((month(readingofmonth_dtp.value) + 1) Mod 12, readtaken_dtp.year + 1)
                        If CDate("" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) Mod 12 & "/" & readtaken_dtp.year + 1 & "") < readtaken_dtp.value Then
                            MsgBox "Reading Date can Not be More than 1 Month"
                            readtaken_dtp.Day = 1
                            Exit Sub
                        End If
                    Else
                        totalday = totaldayinmonth(month(readingofmonth_dtp.value) + 1, readingofmonth_dtp.year)
                        Debug.Print "" & totalday & " / " & (month(readingofmonth_dtp.value) + 1) & " / " & readtaken_dtp.year & ""
                        If CDate("" & totalday & "/" & (month(readingofmonth_dtp.value) + 1) & "/" & readtaken_dtp.year & "") < readtaken_dtp.value Then
                            MsgBox "Reading Date can Not be More than 1 Month"
                            readtaken_dtp.Day = 1
                            Exit Sub
                        End If
                    End If
                End If
                
                
                
                For i = 1 To readinggrid.Rows - 1
                    If readinggrid.TextMatrix(i, 2) = "" Then
                        MsgBox "Please Fill All Readings"
                        Exit Sub
                    End If
                Next
                
                For i = 1 To readinggrid.Rows - 1
                    qry = "update reading_t set creading='" & readinggrid.TextMatrix(i, 2) & "',dateofreading= '" & str(readtaken_dtp.value) & "' where rid =" & readinggrid.TextMatrix(i, 6) & ""
                    Call update(qry)
                Next
                MsgBox " Record Updated Successfully", vbInformation
                Call form_reset
        Case 3
    End Select
End Sub

Private Function dupli() As Boolean
Dim i As Integer
For i = 1 To readinggrid.Rows - 1
    If (readinggrid.TextMatrix(i, 0) = Val(ivrs_txt)) Then
    dupli = True
    Exit Function
    End If
Next

dupli = False

End Function


Private Sub form_reset()
        readinggrid.FixedCols = 0
        readinggrid.Rows = 1
        
        fg_click_state = False
        state = 3
        fg_rowcount = 0
        
       
        name_txt.Text = ""
        creading_txt = ""
        ivrs_txt = ""
        
        creading_txt.Enabled = False
        ivrs_txt.Enabled = False
        ivrssrc_cmd.Enabled = False
        readinggrid.Enabled = False
        
        readingofmonth_dtp.Enabled = False
        readtaken_dtp.Enabled = False
        
        
        del_cmd.Enabled = False
        s_cmd.Enabled = False
        
        Check1.value = 0
        Check1.Visible = False
        Label2.Visible = False
End Sub

Private Sub src_cmd_Click()
reading_src_frm.Show vbModal
End Sub



'///// month last date bill generate allow or niot
