VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form paybill_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "paybill_frm"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker billgen_dtp 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   96403457
      CurrentDate     =   42563
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox payamt_txt 
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker paydate_dtp 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   96403457
      CurrentDate     =   42563
   End
   Begin VB.Label Label4 
      Caption         =   "Bill Paid Date"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Bill generate Date "
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Bill Amount"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Pay Amount"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "paybill_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public month As String

Public year As Integer
Public ivrs As Long
Dim extraamt As Long
Dim rs As ADODB.Recordset
Private Sub OKButton_Click()
    If CDate(paydate_dtp.value) > CDate(billgen_dtp.value) Then
        If payamt_txt <> "" Then
             If Val(Text2) <= Val(payamt_txt) Then
                 Dim str As String
                 Debug.Print "update bill_t set paid=1 where billofmonth<= #1/" & month & "/" & year & "# and ivrs=1 "
                 str = "update bill_t set paid=1 where billofmonth<= #" & month & "/1/" & year & "# and ivrs='" & ivrs & "'"
                 
                 update (str)
                 
                 str = "insert into paybill_t values('" & ivrs & "','" & paydate_dtp.value & "','" & payamt_txt.Text & "')"
                 insert (str)
                 
                 
                 If Val(payamt_txt) > Val(Text2) Then
                    extraamt = Val(payamt_txt) - Val(Text2)
                    Set rs = New ADODB.Recordset
                    rs.Open "select * from connection_t where ivrs='" & ivrs & "'", bms_cn, 3, 3
                    
                    If Not rs.Fields("amtstore") Then
                        extraamt = extraamt + Val(rs.Fields("amtstore"))
                    End If
                    
                    update ("update connection_t set amtstore='" & extraamt & "'where ivrs='" & ivrs & "'")
                 End If
                 
                 
                 
                 MsgBox "Bill payed Successfully"
                
                 Unload Me
            Else
                MsgBox "payment amount can not be less than bill amount", vbInformation
            End If
        Else
            MsgBox "Payment Amount can not be empty", vbInformation
        End If
    Else
        MsgBox "Pay date can not be smaller Than Bill Generation date", vbInformation
    End If
End Sub
