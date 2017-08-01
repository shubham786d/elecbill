VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form billgen_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bill Form"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "billgen_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox year_txt 
      Height          =   285
      Left            =   7680
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton show_cmd 
      Height          =   495
      Left            =   9720
      Picture         =   "billgen_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid billgrid 
      Height          =   4335
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   -2147483624
      BackColorFixed  =   8438015
      GridColor       =   33023
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
   Begin MSComCtl2.DTPicker billgendate_dtp 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   97845249
      CurrentDate     =   42550
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Format          =   97845249
      CurrentDate     =   42550
   End
   Begin MSComCtl2.DTPicker readingofmonth_dtp 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM/yyyy"
      Format          =   97845251
      CurrentDate     =   42514
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Bill Generation date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6255
      Left            =   2760
      Top             =   360
      Width           =   9975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6495
      Left            =   2760
      Top             =   240
      Width           =   9975
   End
End
Attribute VB_Name = "billgen_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public state As Integer
Dim rst As ADODB.Recordset
Dim rcst As ADODB.Recordset ' temp
Dim rcstt As ADODB.Recordset ' temp
Dim Recordsett As Recordset '' temp
Dim Recordset2 As Recordset ' temp
Dim Recordset3 As Recordset ' for grid
Dim fg_rowcount As Integer

Dim minmonth As String


Dim billid As Long
Dim unit As Long
Dim unitrate As Long
Dim totalbill As Long
Dim meterrent As Long
Dim fixcharge As Long
Dim perunittax As Long
Dim percentchg As Long
Dim asloadtax As Long
Dim oldbill As Long
Dim unitcharge As Long
Dim currentbill As Long
Dim recordconsumtion As Long
Dim bsubsity As Long
Dim boardempsubsity As Long
Dim amtstor As Long

Private Sub billgrid_Click()
    readingofmonth_dtp.month = billgrid.Row
    readingofmonth_dtp.Day = 1
    readingofmonth_dtp.year = year_txt.Text
    
    If billgrid.Col = 4 Then
        If billgrid.CellBackColor = vbWhite Then '// checking for existing bill records
            Dim i As Long
            Dim str As String
            Dim readmonth As String
        
        '// checking for existing bill records
            Set rcst = New ADODB.Recordset
            rcst.CursorLocation = adUseClient
            Debug.Print "select * from bill_t where DatePart('m', [billofmonth]) = " & readmonth & " "
            rcst.Open "select * from bill_t where DatePart('m',[billofmonth]) = " & billgrid.TextMatrix(billgrid.Row, 5) & " and DatePart('yyyy',[billofmonth]) = " & year_txt.Text & " ", bms_cn, 3, 3
            
            If rcst.RecordCount > 0 Then
                MsgBox "selected month Bill is already Generated", vbInformation
                Exit Sub
            End If
        
        
        '// checking date should be > admin set date
        Set rcst = New ADODB.Recordset
        rcst.CursorLocation = adUseClient
        rcst.Open "select *  from settings_t ", bms_cn, 3, 3
        
        If Not IsNull(rcst.Fields(15)) Then
            bsubsity = rcst.Fields(15)
        End If
    
        If CDate(rcst.Fields(0)) > readingofmonth_dtp Then
            MsgBox "bill can not be taken from less than starting of  bill date set by ADMIN which is " & rcst.Fields(0) & "", vbInformation
            Exit Sub
        End If
        
        
        '//checking for previous bill generated or not
         Set rcst = New ADODB.Recordset
        rcst.CursorLocation = adUseClient
        rcst.Open "select max(billofmonth) from bill_t ", bms_cn, 3, 3
    
        If rcst.RecordCount > 0 Then
            If Not IsNull(rcst.Fields(0)) Then
                DTPicker1.value = DateAdd("m", 1, CDate(rcst.Fields(0)))
                If DTPicker1.month <> readingofmonth_dtp.month Then
                     MsgBox "Previous Month Bill is Not Generated", vbInformation
                     Exit Sub
                End If
            End If
        End If
    
        
        
        
    '// Bill generation date  >= date of reading
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    Debug.Print "select * from reading_t where  DatePart('m', [readingofmonth])= " & billgrid.TextMatrix(billgrid.Row, 5) & " and DatePart('yyyy', [readingofmonth])=" & year_txt.Text & " "
    rst.Open "select * from reading_t where  DatePart('m', [readingofmonth])= " & billgrid.TextMatrix(billgrid.Row, 5) & " and DatePart('yyyy', [readingofmonth])=" & year_txt.Text & " ", bms_cn, 3, 3
    
    
        If (rst.BOF = False) And (rst.EOF = False) Then
            If CDate(rst.Fields(2)) > billgendate_dtp.value Then
                MsgBox "Bill generation date should be grater than or equal to  reading date which is " & rst.Fields(2), vbInformation
                Exit Sub
            End If
        End If
    
    
    

    If rst.RecordCount > 0 Then
        For i = 0 To rst.RecordCount - 1
            fixcharge = 0
            perunittax = 0
            asloadtax = 0
            oldbill = 0
            
            unit = rst.Fields(5) - rst.Fields(4)
            calmeterrent rst.Fields(1)
            caltax rst.Fields(1)
            Call caloldbill(rst.Fields(1))
            Call findunitrate(rst.Fields(1))
            Call calrecordconsum(rst.Fields(1))
            Call calinterest
            Call calboardempsubsity(rst.Fields(1))
            readingofmonth_dtp.Day = 1
            currentbill = (unit * unitrate) + meterrent + fixcharge + perunittax + asloadtax + Val(boardempsubsity)
            totalbill = currentbill + oldbill
            Call storeamtdeduct(rst.Fields(1))
            str = "insert into bill_t values(" & billid & "," & rst.Fields(0) & "," & rst.Fields(1) & ",'" & readingofmonth_dtp.value & "','" & billgendate_dtp.value & "'," & unit & ",'" & totalbill & "'," & meterrent & "," & fixcharge & ",''," & asloadtax & "," & oldbill & ",'" & DateAdd("d", 15, billgendate_dtp) & "','" & DateAdd("d", 25, billgendate_dtp) & "','" & unit * unitrate & "'," & currentbill & ",0," & recordconsumtion & "," & boardempsubsity & "," & amtstor & ",'" & bms_mdi.user_name_cmd.Caption & "')"
            insert (str)
            billid = billid + 1
            rst.MoveNext
        Next
        loadinggen_frm.Show vbModal
        MsgBox "Bill Generated !!! ", vbInformation
        fixcharge = 0
        perunittax = 0
        state = 3
        'billgendate_dtp.Enabled = False
        readingofmonth_dtp.Enabled = False
        'Option1.Enabled = False
        'Option2.Enabled = False
        show_cmd = True
    Else
        MsgBox "NO Reading Found For This Month", vbInformation
    End If
    End If
    End If
    
End Sub

Private Sub Form_Load()
    billgrid.Cols = 7
    billgrid.Rows = 1
    billgrid.ColWidth(0) = 1600
    billgrid.ColWidth(1) = 1600
    billgrid.ColWidth(2) = 1600
    billgrid.ColWidth(3) = 1600
    billgrid.ColWidth(4) = 1200
    
    billgrid.FixedCols = 0
    fg_rowcount = 0
       
       
    billgrid.TextMatrix(0, 0) = "Month"
    billgrid.TextMatrix(0, 1) = "Reading"
    billgrid.TextMatrix(0, 2) = "User ID"
    billgrid.TextMatrix(0, 3) = "Gen Date"
    billgrid.TextMatrix(0, 4) = "Status"
    'billgrid.TextMatrix(0, 5) = "View"
    'billgrid.TextMatrix(0, 6) = "Status"
    'billgrid.TextMatrix(0, 7) = "Print"
    
    
    Set rst = New ADODB.Recordset
         rst.CursorLocation = adUseClient
            
         rst.Open ("select max(billid) from bill_t"), bms_cn, 3, 3
        ' rst
         
         If rst.RecordCount > 0 Then
              If IsNull(rst.Fields(0)) Then
                  billid = 1
              Else
                  billid = rst.Fields(0) + 1
              End If
         End If
         
         rst.Close

    billgendate_dtp.CustomFormat = "dd/MM/yyyy"
    billgendate_dtp.Format = dtpCustom
    
    
End Sub

Private Sub show_cmd_Click()
    
    Dim months(12) As String
    months(0) = "Jan"
    months(1) = "Feb"
    months(2) = "Mar"
    months(3) = "Apr"
    months(4) = "May"
    months(5) = "Jun"
    months(6) = "Jul"
    months(7) = "Aug"
    months(8) = "Sep"
    months(9) = "Oct"
    months(10) = "Nov"
    months(11) = "Dec"
    
    If year_txt.Text = "" Then
        MsgBox "Please enter year of Bill"
        Exit Sub
    End If
    
   
    
    
    
    Dim i As Long
    fg_rowcount = 0
    billgrid.Rows = 1
       
    For i = 0 To 11
        Set Recordset3 = New ADODB.Recordset
        Recordset3.CursorLocation = adUseClient
        Recordset3.Open ("select * from bill_t where DatePart('m',[billofmonth])=" & i + 1 & " And DatePart('yyyy',[billofmonth]) = '" & year_txt.Text & "'"), bms_cn, 3, 3
        Debug.Print "select * from bill_t where DatePart('m',[billofmonth])=" & i + 1 & " And DatePart('yyyy',[billofmonth]) = '" & year_txt.Text & "'"
            
            fg_rowcount = fg_rowcount + 1
            billgrid.Rows = billgrid.Rows + 1
            billgrid.TextMatrix(i + 1, 0) = months(i)
                
            Set rcstt = New ADODB.Recordset
            rcstt.CursorLocation = adUseClient
            rcstt.Open "select * from bill_t where DatePart('m',[billofmonth]) = " & i + 1 & " and DatePart('yyyy',[billofmonth]) = " & year_txt.Text & " ", bms_cn, 3, 3
            
            
            Set Recordsett = New ADODB.Recordset
            Recordsett.CursorLocation = adUseClient
            Recordsett.Open ("select * from reading_t where DatePart('m',[readingofmonth])=" & i + 1 & " And DatePart('yyyy',[readingofmonth]) = '" & year_txt.Text & "'"), bms_cn, 3, 3
            
            
            If Recordsett.RecordCount <> 0 Then
                billgrid.TextMatrix(i + 1, 1) = "Taken"
                
                If rcstt.RecordCount <> 0 Then
                    billgrid.TextMatrix(i + 1, 2) = rcstt.Fields(20)
                End If
                
                If Recordset3.RecordCount = 0 Then
                    billgrid.TextMatrix(i + 1, 3) = "Not Gen"
                Else
                    billgrid.TextMatrix(i + 1, 3) = Recordset3.Fields(4)
                End If
            Else
                billgrid.TextMatrix(i + 1, 1) = "Not Taken"
            End If
            
            If Recordset3.RecordCount <> 0 Then
                billgrid.Row = i + 1
                billgrid.Col = 4
                Debug.Print App.Path & "\img\button(5).gif"
                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\gened.gif")
                billgrid.CellPictureAlignment = 3
                billgrid.CellBackColor = vbBlack
            Else
                billgrid.Row = i + 1
                billgrid.Col = 4
                Debug.Print App.Path & "\img\button(5).gif"
                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\gen.gif")
                billgrid.CellPictureAlignment = 3
                billgrid.CellBackColor = vbWhite
            End If
            billgrid.TextMatrix(i + 1, 5) = i + 1
'            billgrid.Row = i
'            billgrid.Col = 6
'
'            If Recordset3.Fields("paid") = True Then
'                billgrid.CellBackColor = vbBlack
'                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\button (2).gif")
'            Else
'                billgrid.CellBackColor = vbWhite
'                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\button (3).gif")
'            End If
'
'
'            billgrid.CellPictureAlignment = 3
'
'            billgrid.Row = i
'            billgrid.Col = 7
'            Set billgrid.CellPicture = LoadPicture(App.Path & "\img\print.gif")
'            billgrid.CellPictureAlignment = 3
            
            
            
        Next
        billgrid.ColWidth(5) = 0
        billgrid.ColWidth(6) = 0
End Sub

Private Sub calmeterrent(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient
    Debug.Print "select * from meter_t where mid=(select meter_id from connection_t where ivrs=" & ivrs & ")",
    rcst.Open "select* from metertyp_t where mtid=( select mtypeid from meter_t where mid=(select meter_id from connection_t where ivrs='" & ivrs & "'))", bms_cn, 3, 3
    
    meterrent = rcst.Fields(2)
    
End Sub

Private Sub caltax(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient
    Debug.Print "select * from tariftax_t where tarif_id=(select tarif_id from connection_t where ivr='" & ivrs & "') "
    rcst.Open "select * from tariftax_t where tarif_id=(select tarif_id from connection_t where ivrs='" & ivrs & "') ", bms_cn, 3, 3
    
    Dim i As Long
    For i = 1 To rcst.RecordCount
        Set Recordsett = New ADODB.Recordset
        Recordsett.CursorLocation = adUseClient
        Debug.Print
        Recordsett.Open "select * from taxtype_t where tid=" & rcst.Fields(1) & "", bms_cn, 3, 3
 
        If Recordsett.Fields(2) <> "" Then
            percentchg = percentchg + Recordsett.Fields(2)
        ElseIf Recordsett.Fields(3) <> "" Then
            fixcharge = fixcharge + Recordsett.Fields(3)
        ElseIf Recordsett.Fields(4) <> "" Then
            perunittax = perunittax + (perunittax * unit)
        ElseIf Recordsett.Fields(5) <> "" Then
            Set Recordset2 = New ADODB.Recordset
            Recordset2.CursorLocation = adUseClient
            Recordset2.Open "select * from connection_t where ivrs='" & ivrs & "'", bms_cn, 3, 3
            
            asloadtax = asloadtax + (Recordsett.Fields(5) * Recordset2.Fields(4) / 1000)
        End If
        rcst.MoveNext
    Next
    
End Sub

Private Sub caloldbill(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient
    Debug.Print "select * from  bill_t where billofmonth<#" & readingofmonth_dtp.value & "# and paid=0 and ivrs=" & ivrs & " "
    rcst.Open "select * from  bill_t where billofmonth<#" & readingofmonth_dtp.value & "# and paid=0 and ivrs='" & ivrs & "' ", bms_cn, 3, 3
    Dim i As Long
    If rcst.RecordCount <> 0 Then
        For i = 1 To rcst.RecordCount
            oldbill = oldbill + rcst.Fields(6)
            rcst.MoveNext
        Next
    End If
    
End Sub

Private Sub calduedates(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient
    Debug.Print "select * from  bill_t where billofmonth<#" & readingofmonth_dtp.value & "# and paid=0 and ivrs=" & ivrs & " "
    rcst.Open "select * from  bill_t where billofmonth<#" & readingofmonth_dtp.value & "# and paid=0 and ivrs=" & ivrs & " ", bms_cn, 3, 3
    Dim i As Long
End Sub

Private Sub findunitrate(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient
    Dim i As Long
    
    rcst.Open "select * from tarifsetting_t where tarifid=(select tarif_id from connection_t where ivrs='" & ivrs & "') ", bms_cn, 3, 3
    
    If rcst.RecordCount > 0 Then
        For i = 1 To rcst.RecordCount
            If unit >= rcst.Fields(2) And unit <= rcst.Fields(3) Then
                unitrate = rcst.Fields(4)
                Exit Sub
            End If
            rcst.MoveNext
        Next
        rcst.MoveLast
        unitrate = rcst.Fields(4)
    End If
    
End Sub

Private Sub calrecordconsum(ivrs As Long)
'    Set rcst = New ADODB.Recordset
'    rcst.CursorLocation = adUseClient
'
'    rcst.Open "select * from connection_t where ivrs='" & ivrs & "'", bms_cn, 3, 3
'
'    If rcst.RecordCount > 0 Then
'        recordconsumtion = Val(rcst.Fields("recordedConsumption"))
'    End If

End Sub
Private Sub calinterest()

End Sub


Private Sub calboardempsubsity(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient

    rcst.Open "select * from connection_t where ivrs='" & ivrs & "'", bms_cn, 3, 3
    
    If rcst.RecordCount > 0 Then
        If rcst.Fields("cuntype") = "B" Then
            boardempsubsity = bsubsity
        Else
            boardempsubsity = 0
        End If
    End If

End Sub

Private Sub storeamtdeduct(ivrs As Long)
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient

    rcst.Open "select * from connection_t where ivrs='" & ivrs & "'", bms_cn, 3, 3
     
        
    If rcst.RecordCount > 0 Then
        If Not IsNull(rcst.Fields("amtstore")) Then
            amtstor = Val(rcst.Fields("amtstore"))
            'totalbill = totalbill - amtstor
            'oldbill = oldbill - amtstor
        End If
    End If
End Sub


Private Sub year_txt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case 48 To 57 'numaric
            Case 8      'backspace
            Case Else
              KeyAscii = 0
    End Select
End Sub
