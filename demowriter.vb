Option Explicit

    Dim conn As New ADODB.Connection
    Dim rs As Recordset
    
    


Sub Button1_Click()
    
    Dim Tagname1 As String
    Dim Time1 As String
    Dim Value1 As String
    
    Dim Status As Variant
    
    conn = CreateObject("ADODB.Connection")
    conn.Open ("IP21")
    Sheet1.Range("tag").Select
    Tagname1 = ActiveCell.Value
    
    Sheet1.Range("timestamp").Select
    ActiveCell.Value = Now
    Time1 = ActiveCell.Value
    
    Sheet1.Range("value").Select
    Value1 = ActiveCell.Value
    
    ' FOR INSERTING INTO THE HISTORY REPEAT AREA
    'Status = WriteValAtTime(Tagname1, Time1, Value1)
    
    ' FOR UPDATING THE RECORD'S INPUT FIELDS
    Status = UpdateInputValue(Tagname1, Time1, Value1)
    conn.Close

End Sub


    Private Function WriteValAtTime(vTag, vTime, vValue)
        On Error GoTo errorhandler:
        vV1 = CStr(vValue)
        vT1 = Format(vTime, "yyyy-mmm-dd HH:mm:ss")
        
        Set rs = conn.Execute("INSERT INTO " + vTag + "(IP_TREND_VALUE, IP_TREND_TIME, IP_TREND_QSTATUS) VALUES ('" + vV1 + "',timestamp'" + vT1 + "','GOOD')")
        
        Exit Function
errorhandler:
    MsgBox (Err.Description)
    End Function

Private Function UpdateInputValue(vTag, vTime, vValue)
        Dim vV1 As String
        Dim vT1 As String
        
        On Error GoTo errorhandler:
        vV1 = CStr(vValue)
        vT1 = Format(vTime, "yyyy-mmm-dd HH:mm:ss")
        
        Dim query As String
        query = "UPDATE " + vTag + " SET IP_INPUT_VALUE ='" + vV1 + "', IP_INPUT_TIME = timestamp'" + vT1 + "', QSTATUS(IP_INPUT_VALUE) = 'GOOD'" 
        
        Set rs = conn.Execute(query)
        
        Exit Function
errorhandler:
    MsgBox (Err.Description)
    End Function


