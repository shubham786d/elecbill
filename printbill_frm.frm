VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form printbill_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton printall_cmd 
      Caption         =   "Print All"
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton show_cmd 
      Caption         =   "show"
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox month_cmb 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox year_txt 
      Height          =   285
      Left            =   6600
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton ivrssrc_cmd 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8520
      TabIndex        =   1
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox ivrs_txt 
      Height          =   285
      Left            =   6600
      MaxLength       =   255
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid billgrid 
      Height          =   3375
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Month"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Year"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "printbill_frm"
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
    If billgrid.Col = 5 Then
        
End Sub

Private Sub Form_Load()
     billgrid.Cols = 7
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

Private Sub ivrssrc_cmd_Click()
    con_src_frm.cldform = 4
    con_src_frm.Show vbModal
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
                Set billgrid.CellPicture = LoadPicture(App.Path & "\img\print.gif")
                billgrid.CellPictureAlignment = 3
                billgrid.Row = i
                billgrid.Col = 6
                
                
                billgrid.CellPictureAlignment = 3
                rst_bill.MoveNext
                
            Next
    Else
        fg_rowcount = 0
        billgrid.Rows = 1
        MsgBox "Slected Date Bill is Not Genereted"
    End If
End Sub

