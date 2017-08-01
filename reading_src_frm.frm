VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form reading_src_frm 
   Caption         =   "Form3"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form3"
   ScaleHeight     =   5160
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker readingofmonth_dtp 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107741185
      CurrentDate     =   42514
   End
   Begin VB.Label Label3 
      Caption         =   "mobth of reading"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "reading search"
      Height          =   255
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "reading_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_read As Recordset
Dim rsr As Recordset
Dim Recordsett As Recordset

Const Checked = "þ"
Const UnChecked = "q"


Private Sub con_otp_Click()
Label4.Visible = True
Text1.Visible = True
Label4.Caption = "Consumer Ivrs:"
End Sub

Private Sub DTPicker1_CloseUp()

    
End Sub

Private Sub DataGrid1_Click()
           Dim i As Long
           Dim readmonth As String
           With reading_frm
            .fg_rowcount = 0
            .readinggrid.Rows = 1
            
            'If readingofmonth_dtp.month < 10 Then
             '   readmonth = "0" & readingofmonth_dtp.month
            'Else
               readmonth = readingofmonth_dtp.month
            'End If
            
            .readinggrid.Enabled = True
            .creading_txt.Enabled = True
            .ivrs_txt.Enabled = True
            .ivrssrc_cmd.Enabled = True
            Set rs_read = New ADODB.Recordset 'geting ivrs
            Debug.Print "select * from reading_t where readingofmonth like #" & readmonth & "/1/" & readingofmonth_dtp.year & "#"
            rs_read.Open "select * from reading_t where readingofmonth like #" & readmonth & "/1/" & readingofmonth_dtp.year & "#", bms_cn, 3, 3
            
            For i = 1 To rs_read.RecordCount
                 .fg_rowcount = .fg_rowcount + 1
                 .readinggrid.Rows = .readinggrid.Rows + 1
                 .readinggrid.TextMatrix(i, 6) = rs_read.Fields(0)
                 .readinggrid.TextMatrix(i, 0) = rs_read.Fields(1)
                 
                 .readingofmonth_dtp.value = rs_read.Fields(3)
                 .readingofmonth_dtp.Day = 1
                 .readinggrid.TextMatrix(i, 1) = rs_read.Fields(4)
                 .readinggrid.TextMatrix(i, 2) = rs_read.Fields(5)
                 .readtaken_dtp = rs_read.Fields(2)
                 
                 .readinggrid.Row = i
                 .readinggrid.Col = 3
                 .readinggrid.CellFontName = "Wingdings"
                 .readinggrid.CellFontSize = 14
                 .readinggrid.CellAlignment = flexAlignCenterCenter
                 .readinggrid.Text = UnChecked

                 Set Recordsett = New ADODB.Recordset ' geting min unit
                 Recordsett.Open "select * from tarif_t where tarifid =(select tarif_id from connection_t where ivrs='" & rs_read.Fields(1) & "')  ", bms_cn, 3, 3
                 
                 .readinggrid.TextMatrix(i, 4) = Recordsett.Fields(4)
                 
                 
                 .readinggrid.Row = i
                 .readinggrid.Col = 5
                 .readinggrid.CellFontName = "Wingdings"
                 .readinggrid.CellFontSize = 14
                 .readinggrid.CellAlignment = flexAlignCenterCenter

                 If rs_read.Fields(6) = True Then  'seting billed
                  .readinggrid.TextMatrix(i, 6) = Checked
                 Else
                  .readinggrid.TextMatrix(i, 5) = UnChecked
                 End If
                 rs_read.MoveNext
             Next
             .s_cmd.Enabled = True
             .state = 2
             .readtaken_dtp.Enabled = True
             .readingofmonth_dtp.Enabled = False
            End With
            Unload Me
End Sub

Private Sub month_otp_Click()
Label4.Visible = False
Text1.Visible = False
End Sub

Private Sub readingofmonth_dtp_CloseUp()
Set rs_read = New ADODB.Recordset
    Dim readmonth As String

'   ' If readingofmonth_dtp.month < 10 Then
'       readmonth = "0" & readingofmonth_dtp.month
'    Else
      readmonth = readingofmonth_dtp.month
    'End If

    rs_read.CursorLocation = adUseClient
    Debug.Print "select * from reading_t where readingofmonth like '" & readmonth & "/%/" & readingofmonth_dtp.year & "'"
    rs_read.Open "select * from reading_t where readingofmonth like '" & readmonth & "/%/" & readingofmonth_dtp.year & "'", bms_cn, 3, 3
    
    If rs_read.RecordCount <> 0 Then
        Set DataGrid1.DataSource = rs_read
    Else
        Set DataGrid1.DataSource = Nothing
    End If
    
End Sub
