Attribute VB_Name = "Module1"
Public Function JalankanSQL(sql As String) As ADODB.Recordset
On Error GoTo Messages

Dim AC As New ADODB.Connection

If AC.State = adStateOpen Then AC.Close
Set AC = Nothing

AC.CursorLocation = adUseClient
AC.Properties.Refresh

AC.Open ("DSN=DSNProduksi")

Set JalankanSQL = AC.Execute(sql)
Exit Function

Messages:
MsgBox "KONEKSI KE SERVER ERROR!", vbCritical + vbOKOnly, "Messages"
End
End Function

