Attribute VB_Name = "Module1"
Public bms_cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

Private Sub main()
    
    Set bms_cn = New ADODB.Connection
    
    bms_cn.ConnectionString = "provider=microsoft.ace.oledb.12.0;data source=" & App.Path & "\database\BMS.accdb"
    
    bms_cn.Open
    
    If bms_cn.state = 1 Then
        loading_frm.Show
        'bms_mdi.Show
    Else
        MsgBox "database is not connected properly"
    End If
    
End Sub


Public Sub setcombo(query As String, combo_cmb As ComboBox, txt As String, frontfld As Integer, bckfld As Integer)
        Set rs = New ADODB.Recordset
        
        rs.Open (query), bms_cn, 3, 3
        
        Dim i As Integer
        combo_cmb.Clear
        combo_cmb.AddItem txt
        combo_cmb.ListIndex = 0
        
        For i = 1 To rs.RecordCount
            combo_cmb.AddItem rs.Fields(frontfld)
            combo_cmb.ItemData(i) = rs.Fields(bckfld)
            rs.MoveNext
        Next
        
        combo_cmb.ListIndex = 0
End Sub

Public Sub insert(query As String)
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.CommandText = query
        Debug.Print cmd.CommandText
        cmd.ActiveConnection = bms_cn
        cmd.Execute
End Sub

Public Sub delete(query As String)
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.CommandText = query
        Debug.Print cmd.CommandText
        cmd.ActiveConnection = bms_cn
        cmd.Execute
End Sub

Public Sub update(query As String)
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdText
        cmd.CommandText = query
        Debug.Print cmd.CommandText
        cmd.ActiveConnection = bms_cn
        cmd.Execute
End Sub

Public Sub updatecombo(cmb As ComboBox, value As Integer)
        For X = 0 To cmb.ListCount
            If cmb.ItemData(X) = value Then
                cmb.ListIndex = X
                Exit For
            End If
        Next
        
     
End Sub

Public Sub form_borderset(form_frame As Frame)
form_frame.Move Screen.ActiveForm.Left + Screen.ActiveForm.Width / 2 - form_frame.Width / 2, Screen.ActiveForm.Top + 250
End Sub

