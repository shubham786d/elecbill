VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tariftax_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "tariftax_frm.frx":0000
   ScaleHeight     =   9855
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton src_cmd 
      Caption         =   "Search"
      Height          =   735
      Left            =   9840
      Picture         =   "tariftax_frm.frx":14DB7
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   720
      Width           =   735
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
      Left            =   4440
      Picture         =   "tariftax_frm.frx":15429
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3240
      Width           =   1215
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
      Left            =   5880
      Picture         =   "tariftax_frm.frx":158DE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   1455
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
      Left            =   7440
      Picture         =   "tariftax_frm.frx":1607E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox taxname_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   18
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton taxtypsrc_cmd 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox taxtype_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox taxvalue_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   15
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton co_ex_cmd 
      Height          =   375
      Left            =   8400
      Picture         =   "tariftax_frm.frx":16922
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton new_cmd 
      Height          =   375
      Left            =   4080
      Picture         =   "tariftax_frm.frx":16E10
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton del_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      Picture         =   "tariftax_frm.frx":1733D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton s_cmd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      Picture         =   "tariftax_frm.frx":1790A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton tarifsrc_cmd 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox purpose_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox phase_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox ctype_txt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid taxgrid 
      Height          =   2175
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3836
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
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tax Type Name"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ta"
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phase given"
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tarif Type"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "connection type"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purpose"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   6375
      Left            =   1680
      Top             =   480
      Width           =   9615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   1680
      Top             =   360
      Width           =   9615
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "tariftax_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public taxid As Long
Public tarif_id As Long
Public fg_rowcount As Integer
Dim fg_click_state As Boolean
Dim state As Integer

Private Type fg_updatestate
     tax_id As Long
     taxname As String
     datastate As Integer         '0=normal ; 1= inserted; 2=updated ; 3= delete
End Type


Private back_matrix() As fg_updatestate
Public backm_count As Integer


Private Sub add_cmd_Click()
    If taxname_txt.Text <> "" Then
        If isdup() = True Then
            fg_rowcount = fg_rowcount + 1
            taxgrid.Rows = taxgrid.Rows + 1
        
            taxgrid.TextMatrix(fg_rowcount, 1) = taxname_txt.Text
            taxgrid.TextMatrix(fg_rowcount, 0) = fg_rowcount
            taxgrid.TextMatrix(fg_rowcount, 2) = taxtype_txt.Text
            taxgrid.TextMatrix(fg_rowcount, 3) = taxvalue_txt.Text
            
            
            ReDim Preserve back_matrix(backm_count)
            
            back_matrix(backm_count).taxname = taxname_txt
            back_matrix(backm_count).tax_id = taxid
            back_matrix(backm_count).datastate = 1
            backm_count = backm_count + 1
        Else
        MsgBox "Selected tax type already exisist", vbInformation
        End If
    Else
        MsgBox "Please search tax type", vbInformation
    End If
End Sub

Private Sub Command1_Click()
Dim i As Integer
        For i = 0 To UBound(back_matrix)
        Form1.MSFlexGrid1.TextMatrix(i, 0) = back_matrix(i).tax_id
        Form1.MSFlexGrid1.TextMatrix(i, 1) = back_matrix(i).taxname
        Form1.MSFlexGrid1.TextMatrix(i, 2) = back_matrix(i).datastate
        Form1.MSFlexGrid1.Rows = Form1.MSFlexGrid1.Rows + 1
        Next
        Form1.Show
End Sub

Private Sub Form_Load()
        ReDim back_matrix(0)
       fg_rowcount = 0
       backm_count = 0
       taxgrid.FixedCols = 0
       

       taxgrid.Cols = 4
       taxgrid.Rows = 1
       
       taxgrid.TextMatrix(fg_rowcount, 1) = "Tax Type Name"
       taxgrid.TextMatrix(fg_rowcount, 0) = "S.no"
       taxgrid.TextMatrix(fg_rowcount, 2) = "Type"
       taxgrid.TextMatrix(fg_rowcount, 3) = "Value"
    
   
End Sub

Private Sub new_cmd_Click()
    s_cmd.Enabled = True
    add_cmd.Enabled = True
    rem_cmd.Enabled = True
    rem_all_txt.Enabled = True
    taxgrid.Enabled = True
    tarifsrc_cmd.Enabled = True
    taxtypsrc_cmd.Enabled = True
    state = 1
End Sub

Private Sub rem_cmd_Click()
    If taxgrid.Rows > 1 Then
        If fg_click_state = True Then
            Dim i, k As Long
            For i = taxgrid.Row To taxgrid.Rows - 2    ' for flax remove display
                taxgrid.TextMatrix(i, 1) = taxgrid.TextMatrix(i + 1, 1)
                taxgrid.TextMatrix(i, 2) = taxgrid.TextMatrix(i + 1, 2)
                taxgrid.TextMatrix(i, 3) = taxgrid.TextMatrix(i + 1, 3)
            Next
            
            For k = 0 To UBound(back_matrix)            ' for back remove
                If taxgrid.TextMatrix(taxgrid.Row, 1) = back_matrix(k).taxname And back_matrix(k).datastate <> 3 Then
                    back_matrix(k).datastate = 3 'removed marked
                    Exit For
                End If
            Next
            
            taxgrid.Rows = taxgrid.Rows - 1
            fg_rowcount = fg_rowcount - 1
        Else
         MsgBox "Please Select Any Record From Table For Removing", vbInformation
        End If
    End If
    
    fg_click_state = False
End Sub

Private Sub s_cmd_Click()
If ctype_txt.Text <> "" Then
    If taxname_txt.Text <> "" Then
        Select Case state
            Case 1
                Dim qry As String
                Dim i As Long
                For i = 0 To UBound(back_matrix)
                    If (back_matrix(i).datastate = 1) Then
                        qry = "insert into tariftax_t values('" & tarif_id & "','" & back_matrix(i).tax_id & "')"
                        Call insert(qry)
                    End If
                Next
                
                s_cmd.Enabled = False
                state = 3
                ReDim back_matrix(0)
                fg_rowcount = 0
                backm_count = 0
                taxgrid.Rows = 1
            Case 2
            Case 3
        End Select
    Else
        MsgBox "Please select tax type", vbInformation
    End If
Else
    MsgBox "Please select tariff", vbInformation
End If
End Sub

Private Sub src_cmd_Click()
tarif_src_frm.callvalue = 3
tarif_src_frm.Show
End Sub

Private Sub tarifsrc_cmd_Click()
tarif_src_frm.callvalue = 2
tarif_src_frm.Show vbModal
End Sub


Private Sub taxgrid_Click()
    If taxgrid.Row > 0 Then
        fg_click_state = True
    End If
End Sub

Private Sub taxtypsrc_cmd_Click()
taxtyp_src_frm.callvalue = 2
taxtyp_src_frm.Show vbModal
End Sub

Private Function isdup() As Boolean
 Dim i As Long
 For i = 1 To taxgrid.Rows - 1
    If (taxname_txt.Text = taxgrid.TextMatrix(i, 1)) Then
        isdup = False
        Exit Function
    End If
 Next
isdup = True
End Function

