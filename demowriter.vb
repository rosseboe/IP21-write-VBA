Private conn As New ADODB.Connection
Private rs As Recordset


Private Sub

    conn = CreateObject("ADODB.Connection")
    conn.Open ("IP21") ' ODBC DSN 

    Tagname1 = "IP21TAGNAME"
    
    Time1 = "21-MAY-19 10:21:12"
    
    Value1 = "2019"
    
    Status = WriteValAtTime(Tagname1, Time1, Value1)
    
    conn.Close

End Sub


Private Function WriteValAtTime(vTag, vTime, vValue)
  On Error GoTo errorhandler:
  vV1 = CStr(vValue)
  vT1 = Format(vTime, "dd-mmm-yy hh:mm:ss")

  Set rs = conn.Execute("INSERT INTO " + vTag + "(IP_TREND_VALUE, IP_TREND_TIME, IP_TREND_QSTATUS) VALUES ('" + vV1 + "','" + vT1 + "','GOOD')")

  Exit Function
  errorhandler:
    MsgBox (Err.Description)
End Function
    
    
