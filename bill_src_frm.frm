VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bill_src_frm 
   Caption         =   "Form3"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form3"
   ScaleHeight     =   5100
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "ALL"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ONE"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox ivrs_txt 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid search_dg 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   8415
      _ExtentX        =   14843
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
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Format          =   95944705
      CurrentDate     =   42514
   End
   Begin VB.Label Label2 
      Caption         =   "IVRS:"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   " bill month "
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "bill_src_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()
Call Option1_Click
Option1.value = True
End Sub



Private Sub Option1_Click()
ivrs_txt.Visible = False
Label2.Visible = False
End Sub

Private Sub Option2_Click()
ivrs_txt.Visible = True
Label2.Visible = True
End Sub


Private Sub readingofmonth_dtp_Change()

            Set rst = New ADODB.Recordset
            rst.CursorLocation = adUseClient
            
            Dim readmonth As String
    
    
            If readingofmonth_dtp.month < 10 Then
               readmonth = "0" & readingofmonth_dtp.month
            Else
               readmonth = readingofmonth_dtp.month
            End If
            
             rst.Open "SELECT distinct  readingofmonth  from reading_t where readingofmonth like '%/" & readmonth & "/" & readingofmonth_dtp.year & "' ", bms_cn, 3, 3
              
        If rst.RecordCount > 0 Then
                     
             Set search_dg.DataSource = rst
             'search_dg.Columns(0).Visible = True
             search_dg.Columns(0).Caption = "Month"
             'search_dg.Columns(2).Caption = "Round Description"
        Else
                Set search_dg.DataSource = Nothing
                
        End If
End Sub

Private Sub search_dg_Click()
        If Option1.value = True Then
            Dim i As Integer
            If search_dg.Row <> -1 Then
                i = search_dg.Row
                search_dg.RowBookmark (i)
                 bill_frm.readingofmonth_dtp = search_dg.Columns(0)
                 bill_frm.readingofmonth_dtp.Enabled = False
                 bill_frm.Option1.value = Option1.value
                 bill_frm.Option2.value = Option2.value
                 bill_frm.generate_cmd.Enabled = True
                 bill_frm.del_cmd.Enabled = True
                 bill_frm.s_cmd.Enabled = True
                 bill_frm.state = 2
                 Unload Me
            End If
        End If
End Sub
