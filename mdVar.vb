Imports ProcessData
Module mdVar
    Public Const Max_Client As Integer = 10
    Public DBServer As String
    Public ConnectionString As String
    Public DataSource As String
    Public IpHost As String
    Public UserName As String
    Public Password As String
    Public objSip2Patron(Max_Client) As clsSip2Patron
    Public blnConnectEnable As Boolean
    Public Sub DisposeAllSip2Patron()
        Dim i As Integer
        For i = 0 To Max_Client
            If Not objSip2Patron(i) Is Nothing Then
                objSip2Patron(i).Dispose(True)
                objSip2Patron(i) = Nothing
            End If
        Next
    End Sub
End Module
