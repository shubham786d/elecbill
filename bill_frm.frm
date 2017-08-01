VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form bill_frm 
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
   Picture         =   "bill_frm.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker billgendate_dtp 
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   96141313
      CurrentDate     =   42550
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Format          =   96141313
      CurrentDate     =   42550
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   9600
      Picture         =   "bill_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   5400
      Picture         =   "bill_frm.frx":152A5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      Picture         =   "bill_frm.frx":157D2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      Picture         =   "bill_frm.frx":15D9F
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   11280
      Picture         =   "bill_frm.frx":164B3
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker readingofmonth_dtp 
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "MMMM/yyyy"
      Format          =   96141315
      CurrentDate     =   42514
   End
   Begin VB.TextBox ivrs_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   4
      Top             =   2880
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "one"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "all"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton generate_cmd 
      BackColor       =   &H000080FF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      Picture         =   "bill_frm.frx":16B25
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " bill Generation date"
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " bill of month "
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IVRS:"
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5655
      Left            =   3960
      Top             =   360
      Width           =   8295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   5895
      Left            =   3960
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "bill_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public state As Integer
Dim rst As ADODB.Recordset
Dim rcst As ADODB.Recordset ' temp
Dim Recordsett As Recordset '' temp
Dim Recordset2 As Recordset ' temp

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
'/////

Private Sub Command1_Click()

 Set rcst = New ADODB.Recordset
 Debug.Print "SELECT *FROM consumer_t, connection_t,bill_t WHERE billofmonth like '1/01/2016' and consumer_t.cid=connection_t.cid and connection_t.ivrs=bill_t.ivrs and connection_t.ivrs='1' "
 rcst.Open "SELECT * From consumer_t, connection_t, bill_t, reading_t, reader_t, tarif_t, perpose_t,settings_t WHERE billofmonth like '1/01/2016' and consumer_t.cid=connection_t.cid and connection_t.ivrs=bill_t.ivrs and reading_t.rid=bill_t.readingid and readerid=reader_t.rid and connection_t.tarif_id=tarif_t.tarifid and tarif_t.perposid=perpose_t.id and  connection_t.ivrs='1';", bms_cn, 3, 3
 
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
 
 DataReport2.Sections("Section1").Controls("totalcon_txt").DataField = rcst.Fields("duedatebycash").Name
 DataReport2.Sections("Section1").Controls("recordcon_txt").DataField = rcst.Fields("duedatebycash").Name
 DataReport2.Sections("Section1").Controls("metercon_txt").DataField = rcst.Fields("duedatebycash").Name
 DataReport2.Sections("Section1").Controls("preread_txt").DataField = rcst.Fields("preading").Name
 DataReport2.Sections("Section1").Controls("curread_txt").DataField = rcst.Fields("creading").Name
 
 DataReport2.Sections("Section1").Controls("lastdatechk2_txt").DataField = rcst.Fields("duedatebychk").Name
 DataReport2.Sections("Section1").Controls("lastdatecash2_txt").DataField = rcst.Fields("duedatebycash").Name
 DataReport2.Sections("Section1").Controls("billmonth_txt").DataField = rcst.Fields("duedatebycash").Name
 DataReport2.Sections("Section1").Controls("ivrsno_txt").DataField = rcst.Fields("connection_t.ivrs").Name
 DataReport2.Sections("Section1").Controls("address_txt").DataField = rcst.Fields("connection_t.address").Name
 DataReport2.Sections("Section1").Controls("billid2_txt").DataField = rcst.Fields("billid").Name
 
 DataReport2.Sections("Section1").Controls("readername_txt").DataField = rcst.Fields("rname").Name
 
 DataReport2.Sections("Section1").Controls("purpose_txt").DataField = rcst.Fields("pname").Name
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
 DataReport2.Sections("Section1").Controls("pendingamt_txt").Caption = rcst.Fields("MeterRent")
 DataReport2.Sections("Section1").Controls("totalamt_txt").Caption = rcst.Fields("totalbill")
 DataReport2.Sections("Section1").Controls("totalafterdate_txt").Caption = rcst.Fields("totalbill")
 
 DataReport2.Show
End Sub

Private Sub Command2_Click()
DataReport1.Orientation = rptOrientLandscape
DataEnvironment1.rsCommand1.Open
Set rcst = New ADODB.Recordset
rcst.DataSource = "select *  from conection_t where ivrs=1"

DataReport2.Show
End Sub

Private Sub del_cmd_Click()
    If state = 2 Then
        If Option1.value = True Then
            
            Dim test As Integer
            test = MsgBox("Do U Want To Delete This Month Bill Records ?", vbYesNoCancel + vbQuestion, "Information")
             
             If test = 6 Then
                Dim str As String
                Dim readmonth As String
            
                If readingofmonth_dtp.month < 10 Then
                   readmonth = "0" & readingofmonth_dtp.month
                Else
                   readmonth = readingofmonth_dtp.month
                End If
                
                str = "delete * from bill_t where  billofmonth like '%/" & readmonth & "/" & readingofmonth_dtp.year & "' "
                Call delete(str)
                MsgBox "This Month Bill Records deleted successfully", vbInformation
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
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
    
    Call Option1_Click
    Option1.value = True


End Sub

Private Sub generate_cmd_Click()
    Dim i As Long
    Dim str As String
    Dim readmonth As String

'    If readingofmonth_dtp.month < 10 Then
'       readmonth = "0" & readingofmonth_dtp.month
'    Else
       readmonth = readingofmonth_dtp.month
' End If
    
    '// checking for existing bill records
    Set rcst = New ADODB.Recordset
    rcst.CursorLocation = adUseClient
    Debug.Print "select * from bill_t where DatePart('m', [billofmonth]) = " & readmonth & " "
    rcst.Open "select * from bill_t where DatePart('m',[billofmonth]) = " & readmonth & " and DatePart('yyyy',[billofmonth]) = " & readingofmonth_dtp.year & " ", bms_cn, 3, 3
    
    If rcst.RecordCount > 0 Then
        MsgBox "selected month Bill is already Generated", vbInformation
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
                 MsgBox "Previous Month Bill is Not Renerated", vbInformation
                 Exit Sub
            End If
        End If
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
    
    '// cheking difference between both date
'    DTPicker1.value = readingofmonth_dtp.value
'    DTPicker1.day = 1
'    Dim day As Integer
'    day = billgendate_dtp.day
'    billgendate_dtp.day = 1
'    If DateAdd("m", 2, DTPicker1.value) <> billgendate_dtp.value Then
'        MsgBox "Bill generation date can not be 2 month More then month of Bill"
'        Exit Sub
'    End If
'    billgendate_dtp.day = day

    '//checking for pevious bill gen date vs new gen date
'    Set rcst = New ADODB.Recordset
'    rcst.CursorLocation = adUseClient
'    rcst.Open "select *  from settings_t ", bms_cn, 3, 3
    
    
    


    
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    Debug.Print "select * from reading_t where  DatePart('m', [readingofmonth])= " & readmonth & " and DatePart('yyyy', [billofmonth])=" & readingofmonth_dtp.year & " "
    rst.Open "select * from reading_t where  DatePart('m', [readingofmonth])= " & readmonth & " and DatePart('yyyy', [readingofmonth])=" & readingofmonth_dtp.year & " ", bms_cn, 3, 3
    
    If (IsEmpty(rst.Fields(2))) Then
        If rst.Fields(2) < readingofmonth_dtp.value Then
            MsgBox "Bill generation date should be grater than reading date"
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
            str = "insert into bill_t values(" & billid & "," & rst.Fields(0) & "," & rst.Fields(1) & ",'" & readingofmonth_dtp.value & "','" & billgendate_dtp.value & "'," & unit & ",'" & totalbill & "'," & meterrent & "," & fixcharge & ",''," & asloadtax & "," & oldbill & ",'" & DateAdd("d", 15, billgendate_dtp) & "','" & DateAdd("d", 25, billgendate_dtp) & "','" & unit * unitrate & "'," & currentbill & ",0," & recordconsumtion & "," & boardempsubsity & "," & amtstor & ")"
            insert (str)
            billid = billid + 1
            rst.MoveNext
        Next
        loadinggen_frm.Show vbModal
        MsgBox "BILL generated ", vbInformation
        fixcharge = 0
        perunittax = 0
        state = 3
        billgendate_dtp.Enabled = False
        readingofmonth_dtp.Enabled = False
        Option1.Enabled = False
        Option2.Enabled = False
    Else
        MsgBox "NO Reading Found For This Month", vbInformation
    End If
    
    generate_cmd.Enabled = False
End Sub

Private Sub new_cmd_Click()
state = 1
Option1.Enabled = True
Option2.Enabled = True
readingofmonth_dtp.Enabled = True
ivrs_txt.Enabled = True
's_cmd.Enabled = True
generate_cmd.Enabled = True
billgendate_dtp.Enabled = True
End Sub

Private Sub Option1_Click()
ivrs_txt.Visible = False
Label2.Visible = False
End Sub

Private Sub Option2_Click()
ivrs_txt.Visible = True
Label2.Visible = True
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


Private Sub src_cmd_Click()
bill_src_frm.Show vbModal
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



