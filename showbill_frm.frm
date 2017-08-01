VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form showbill_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "show Bill"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "showbill_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton printall_cmd 
      Height          =   495
      Left            =   9840
      Picture         =   "showbill_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox ivrs_txt 
      Height          =   285
      Left            =   7080
      MaxLength       =   255
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton show_cmd 
      Height          =   495
      Left            =   9840
      Picture         =   "showbill_frm.frx":152A5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox year_txt 
      Height          =   285
      Left            =   7080
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox month_cmb 
      Height          =   315
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid billgrid 
      Height          =   3375
      Left            =   3120
      TabIndex        =   0
      Top             =   2760
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5953
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
      Left            =   6240
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
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
      Left            =   6240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Month"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6375
      Left            =   2880
      Top             =   360
      Width           =   9855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   3000
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "showbill_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst_bill As ADODB.Recordset ' main
Dim rcst2 As ADODB.Recordset '  used 4 bill
Dim rcst As ADODB.Recordset ' used 4 bill
Dim fg_rowcount As Integer
Dim month As String
Private Sub billgrid_Click()
    Set rcst = New ADODB.Recordset
    If billgrid.Col = 5 Or billgrid.Col = 7 Then
  
         
         Debug.Print "SELECT *FROM consumer_t, connection_t,bill_t WHERE billofmonth like '1/01/2016' and consumer_t.cid=connection_t.cid and connection_t.ivrs=bill_t.ivrs and connection_t.ivrs='1' "
         rcst.Open "SELECT * From consumer_t, connection_t, bill_t, reading_t, reader_t, tarif_t, perpose_t,settings_t,phase_t WHERE DatePart('m',[billofmonth]) = '" & month & "' and DatePart('yyyy',[billofmonth]) ='" & year_txt & "' and consumer_t.cid=connection_t.cid and connection_t.ivrs=bill_t.ivrs and reading_t.rid=bill_t.readingid and readerid=reader_t.rid and connection_t.tarif_id=tarif_t.tarifid and tarif_t.perposid=perpose_t.id and phase_t.pid=tarif_t.phaseid  and  connection_t.ivrs='" & billgrid.TextMatrix(billgrid.Row, 1) & "'", bms_cn, 3, 3
         
         If rcst.RecordCount > 0 Then
        ' MsgBox rcst.RecordCount & rcst.Fields(1).Name
         Set DataReport2.DataSource = rcst
        
         DataReport2.Sections("Section1").Controls("ivrs_txt").DataField = rcst.Fields("connection_t.ivrs").Name
         DataReport2.Sections("Section1").Controls("name_txt").DataField = rcst.Fields("cname").Name
         DataReport2.Sections("Section1").Controls("add_txt").DataField = rcst.Fields("connection_t.address").Name
         DataReport2.Sections("Section1").Controls("mob_txt").DataField = rcst.Fields("mobno").Name
         
         DataReport2.Sections("Section1").Controls("billdate_txt").DataField = rcst.Fields("billgendate").Name
         DataReport2.Sections("Section1").Controls("billnum_txt").DataField = rcst.Fields("billid").Name
         
         DataReport2.Sections("Section1").Controls("meterno_txt").DataField = rcst.Fields("meter_id").Name
         DataReport2.Sections("Section1").Controls("load_txt").DataField = rcst.Fields("Load").Name
         DataReport2.Sections("Section1").Controls("tarif_txt").DataField = rcst.Fields("tarif_id").Name
         DataReport2.Sections("Section1").Controls("secamt_txt").DataField = rcst.Fields("secuamt").Name
         DataReport2.Sections("Section1").Controls("name2_txt").DataField = rcst.Fields("cname").Name
         
         DataReport2.Sections("Section1").Controls("totalunit_txt").DataField = rcst.Fields("totalunit").Name
         DataReport2.Sections("Section1").Controls("amount_txt").DataField = rcst.Fields("totalbill").Name
         DataReport2.Sections("Section1").Controls("monthofread_txt").DataField = rcst.Fields("billofmonth").Name
         
         DataReport2.Sections("Section1").Controls("lastdatechk_txt").DataField = rcst.Fields("duedatebychk").Name
         DataReport2.Sections("Section1").Controls("lastdatecash_txt").DataField = rcst.Fields("duedatebycash").Name
         
         DataReport2.Sections("Section1").Controls("totalcon_txt").DataField = rcst.Fields("totalunit").Name
         DataReport2.Sections("Section1").Controls("recordedcon_txt").Caption = 0
         DataReport2.Sections("Section1").Controls("metercon_txt").DataField = rcst.Fields("totalunit").Name
         DataReport2.Sections("Section1").Controls("preread_txt").DataField = rcst.Fields("preading").Name
         DataReport2.Sections("Section1").Controls("curread_txt").DataField = rcst.Fields("creading").Name
         
         DataReport2.Sections("Section1").Controls("lastdatechk2_txt").DataField = rcst.Fields("duedatebychk").Name
         DataReport2.Sections("Section1").Controls("lastdatecash2_txt").DataField = rcst.Fields("duedatebycash").Name
         DataReport2.Sections("Section1").Controls("billmonth_txt").DataField = rcst.Fields("duedatebycash").Name
         DataReport2.Sections("Section1").Controls("ivrsno_txt").DataField = rcst.Fields("connection_t.ivrs").Name
         DataReport2.Sections("Section1").Controls("address_txt").DataField = rcst.Fields("connection_t.address").Name
         DataReport2.Sections("Section1").Controls("billid2_txt").DataField = rcst.Fields("billid").Name
         
         DataReport2.Sections("Section1").Controls("readername_txt").DataField = rcst.Fields("rname").Name
         
         DataReport2.Sections("Section1").Controls("purpose_txt").DataField = rcst.Fields("perpose_t.pname").Name
         DataReport2.Sections("Section1").Controls("copname1_txt").DataField = rcst.Fields("name1").Name
         DataReport2.Sections("Section1").Controls("copname2_txt").DataField = rcst.Fields("name2").Name
         DataReport2.Sections("Section1").Controls("copnum1_txt").DataField = rcst.Fields("compphnenum1").Name
         DataReport2.Sections("Section1").Controls("copnum2_txt").DataField = rcst.Fields("compphnenum2").Name
         
         
         DataReport2.Sections("Section1").Controls("energyc_txt").Caption = rcst.Fields("unitcharge")
         DataReport2.Sections("Section1").Controls("fixedc_txt").Caption = rcst.Fields("fixedcharge")
         DataReport2.Sections("Section1").Controls("elecc_txt").Caption = rcst.Fields("perunittax")
         DataReport2.Sections("Section1").Controls("loadtax_txt").Caption = rcst.Fields("asloadtax")
         DataReport2.Sections("Section1").Controls("meterrent_txt").Caption = rcst.Fields("MeterRent")
         DataReport2.Sections("Section1").Controls("currentamt_txt").Caption = rcst.Fields("currentbill")
         DataReport2.Sections("Section1").Controls("pendingamt_txt").Caption = rcst.Fields("oldbill")
         DataReport2.Sections("Section1").Controls("totalamt_txt").Caption = rcst.Fields("totalbill")
         DataReport2.Sections("Section1").Controls("totalafterdate_txt").Caption = rcst.Fields("totalbill")
         
         DataReport2.Sections("Section1").Controls("email_txt").Caption = rcst.Fields("emailid")
         DataReport2.Sections("Section1").Controls("curreaddate_txt").Caption = rcst.Fields("dateofreading")
         DataReport2.Sections("Section1").Controls("phase_txt").Caption = rcst.Fields("phase_t.pname")
          DataReport2.Sections("Section1").Controls("totalatend_txt").Caption = rcst.Fields("totalbill")
         DataReport2.Sections("Section1").Controls("phno_txt").Caption = rcst.Fields("phno")
      '   DataReport2.Sections("Section1").Controls("bsubsity_txt").Caption = rcst.Fields("bill_t.boardempsubsity")
         
         Set rcst2 = New ADODB.Recordset
         rcst2.Open "select top 6 * from reading_t where ivrs=" & billgrid.TextMatrix(billgrid.Row, 1) & " order by readingofmonth desc", bms_cn, 3, 3
         
         rcst2.MoveLast
         
         If rcst2.RecordCount > 0 Then
            DataReport2.Sections("Section1").Controls("pmdate1_txt").Caption = rcst2.Fields("readingofmonth")
            DataReport2.Sections("Section1").Controls("readdate1_txt").Caption = rcst2.Fields("dateofreading")
            DataReport2.Sections("Section1").Controls("reading1_txt").Caption = rcst2.Fields("creading")
            DataReport2.Sections("Section1").Controls("unit1_txt").Caption = Val(rcst2.Fields("creading")) - Val(rcst2.Fields("preading"))
            
            rcst2.MovePrevious
            If rcst2.BOF = True Then
                GoTo way
            End If
            
            DataReport2.Sections("Section1").Controls("pmdate2_txt").Caption = rcst2.Fields("readingofmonth")
            DataReport2.Sections("Section1").Controls("readdate2_txt").Caption = rcst2.Fields("dateofreading")
            DataReport2.Sections("Section1").Controls("reading2_txt").Caption = rcst2.Fields("creading")
            DataReport2.Sections("Section1").Controls("unit2_txt").Caption = Val(rcst2.Fields("creading")) - Val(rcst2.Fields("preading"))
            rcst2.MovePrevious
            If rcst2.BOF = True Then
                GoTo way
            End If
            
            DataReport2.Sections("Section1").Controls("pmdate3_txt").Caption = rcst2.Fields("readingofmonth")
            DataReport2.Sections("Section1").Controls("readdate3_txt").Caption = rcst2.Fields("dateofreading")
            DataReport2.Sections("Section1").Controls("reading3_txt").Caption = rcst2.Fields("creading")
            DataReport2.Sections("Section1").Controls("unit3_txt").Caption = Val(rcst2.Fields("creading")) - Val(rcst2.Fields("preading"))
            rcst2.MovePrevious
            If rcst2.BOF = True Then
                GoTo way
            End If
            
            DataReport2.Sections("Section1").Controls("pmdate4_txt").Caption = rcst2.Fields("readingofmonth")
            DataReport2.Sections("Section1").Controls("readdate4_txt").Caption = rcst2.Fields("dateofreading")
            DataReport2.Sections("Section1").Controls("reading4_txt").Caption = rcst2.Fields("creading")
            DataReport2.Sections("Section1").Controls("unit4_txt").Caption = Val(rcst2.Fields("creading")) - Val(rcst2.Fields("preading"))
            rcst2.MovePrevious
            If rcst2.BOF = True Then
                GoTo way
            End If
            
            DataReport2.Sections("Section1").Controls("readdate5_txt").Caption = rcst2.Fields("dateofreading")
            DataReport2.Sections("Section1").Controls("pmdate5_txt").Caption = rcst2.Fields("readingofmonth")
            DataReport2.Sections("Section1").Controls("reading5_txt").Caption = rcst2.Fields("creading")
            DataReport2.Sections("Section1").Controls("unit5_txt").Caption = Val(rcst2.Fields("creading")) - Val(rcst2.Fields("preading"))
           
            
         End If
way:
         If billgrid.Col = 7 Then
            DataReport2.PrintReport , rptRangeFromTo, 1, 1
         Else
            DataReport2.Show
         End If
         
        End If
  End If
    If billgrid.Col = 6 Then
        If billgrid.CellBackColor = vbWhite Then
             paybill_frm.ivrs = billgrid.TextMatrix(billgrid.Row, 1)
             paybill_frm.month = month
             paybill_frm.year = Val(year_txt)
             paybill_frm.billgen_dtp.value = CDate(billgrid.TextMatrix(billgrid.Row, 2))
             paybill_frm.Text2.Text = billgrid.TextMatrix(billgrid.Row, 4)
             paybill_frm.Show vbModal
             show_cmd = True
        End If
    End If
    
'    If billgrid.Col = 7 Then
'
'        DataReport2.PrintReport
'    End If

End Sub

Private Sub Form_Load()
    billgrid.Cols = 8
    billgrid.Rows = 1
    billgrid.ColWidth(0) = 1050
    billgrid.ColWidth(1) = 1050
    billgrid.ColWidth(2) = 1050
    billgrid.ColWidth(3) = 1050
    
    billgrid.FixedCols = 0
    fg_rowcount = 0
       
       
    billgrid.TextMatrix(0, 0) = "Bill Id"
    billgrid.TextMatrix(0, 1) = "Ivrs"
    billgrid.TextMatrix(0, 2) = "Generation date"
    billgrid.TextMatrix(0, 3) = "Total Unit"
    billgrid.TextMatrix(0, 4) = "Total Bill"
    billgrid.TextMatrix(0, 5) = "View"
    billgrid.TextMatrix(0, 6) = "Status"
    billgrid.TextMatrix(0, 7) = "Print"
       
    month_cmb.AddItem "Please select Month"
    month_cmb.AddItem "Jan"
    month_cmb.AddItem "Feb"
    month_cmb.AddItem "Mar"
    month_cmb.AddItem "Apr"
    month_cmb.AddItem "May"
    month_cmb.AddItem "Jun"
    month_cmb.AddItem "Jul"
    month_cmb.AddItem "Aug"
    month_cmb.AddItem "Sep"
    month_cmb.AddItem "Oct"
    month_cmb.AddItem "Nov"
    month_cmb.AddItem "Dec"
    
    Dim i As Integer
    For i = 0 To 12
        month_cmb.ItemData(i) = i
    Next
    month_cmb.ListIndex = 0

    

End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub ivrs_txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Dim i As Integer
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 1) = ivrs_txt.Text Then
                billgrid.Row = i
                billgrid.RowSel = i
                billgrid.Col = 0
                billgrid.ColSel = billgrid.Cols - 1
                Exit Sub
            End If
        Next
        MsgBox "IVRS Not Found In List"
    End If
End Sub

Private Sub ivrssrc_cmd_Click()
con_src_frm.cldform = 4
con_src_frm.Show vbModal
End Sub

Private Sub printall_cmd_Click()
    Dim i As Integer
    For i = 1 To billgrid.Rows - 1
        billgrid.Col = 7
        billgrid.Row = i
        Call billgrid_Click
    Next
End Sub

Private Sub show_cmd_Click()
   
    If year_txt.Text = "" Then
        MsgBox "Please enter year of Bill"
        Exit Sub
    End If
    
    If month_cmb.ListIndex = 0 Then
        MsgBox "Please Select month of Bill"
        Exit Sub
    End If

'     If month_cmb.ItemData(month_cmb.ListIndex) < 10 Then
'       month = "0" & month_cmb.ItemData(month_cmb.ListIndex)
'    Else
   month = month_cmb.ItemData(month_cmb.ListIndex)
'    End If
'
    
    Set rst_bill = New ADODB.Recordset
    rst_bill.CursorLocation = adUseClient
       
    rst_bill.Open ("select * from bill_t where DatePart('m',[billofmonth]) = '" & month & "' and DatePart('yyyy',[billofmonth]) = '" & year_txt.Text & "'"), bms_cn, 3, 3
    Debug.Print "select * from bill_t where DatePart('m',[billofmonth]) = '" & month & "' and DatePart('yyyy',[billofmonth]) = '" & year_txt.Text & "')"
    If rst_bill.RecordCount > 0 Then
            Dim i As Long
            fg_rowcount = 0
            billgrid.Rows = 1
            
            For i = 1 To rst_bill.RecordCount
                fg_rowcount = fg_rowcount + 1
                billgrid.Rows = billgrid.Rows + 1
                billgrid.TextMatrix(i, 0) = rst_bill.Fields(0)
                billgrid.TextMatrix(i, 1) = rst_bill.Fields(2)
                billgrid.TextMatrix(i, 2) = rst_bill.Fields(4)
                billgrid.TextMatrix(i, 3) = rst_bill.Fields(5)
                billgrid.TextMatrix(i, 4) = rst_bill.Fields(6)
                billgrid.Row = i
                billgrid.Col = 5
                Debug.Print App.Path & " \img\button.gif"
                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\button.gif")
                billgrid.CellPictureAlignment = 3
                billgrid.Row = i
                billgrid.Col = 6
                
                If rst_bill.Fields("paid") = True Then
                    billgrid.CellBackColor = vbBlack
                    Set billgrid.CellPicture = LoadPicture(App.Path & "\img\button (2).gif")
                Else
                    billgrid.CellBackColor = vbWhite
                    Set billgrid.CellPicture = LoadPicture(App.Path & "\img\button (3).gif")
                End If
                
                
                billgrid.CellPictureAlignment = 3
                
                billgrid.Row = i
                billgrid.Col = 7
                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\print.gif")
                billgrid.CellPictureAlignment = 3
                
                rst_bill.MoveNext
                
            Next
    Else
        fg_rowcount = 0
        billgrid.Rows = 1
        MsgBox "Slected Date Bill is Not Genereted"
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
