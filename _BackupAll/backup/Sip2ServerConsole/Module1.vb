Imports WinSocket
Imports System.IO
Imports ProcessData
Imports System.Net
Imports System.Net.Sockets
Module Module1
    Private arrRequest(Max_Client) As String
    Private arrLastResponse(Max_Client) As String
    Private arrClient(Max_Client) As String
    Private strTimeZone As String
    Private strCurrencyUnit As String
    Private strChecksumEnabled As String
    Private intSIP2Port As Integer
    Private WithEvents ServerWinsock As clsServerWinsock
    Private clientSocket As Socket

    Private Const Max_Client As Integer = 10
    Private objSip2Patron(Max_Client) As clsSip2Patron

    Private ConnectionString As String
    Private ConnectEnable As Boolean
    Private DBServer As String
    Sub Main()
        Dim tblInfo As DataTable
        ' Dim objLg As New LibolLogin.clsLibolLogin
        'ConnectionString = objLg.GetConnectionString()
        'DBServer = objLg.DBServer
        'tblInfo = GetXmlFile(a.StartupPath & "\Sip2Server.xml")
        'ConnectionString = tblInfo.Rows(0).Item("ConnectString")
        'DBServer = "SQLSERVER"
        'strTimeZone = tblInfo.Rows(0).Item("TimeZone")
        'strCurrencyUnit = tblInfo.Rows(0).Item("Currency")
        'strChecksumEnabled = tblInfo.Rows(0).Item("Checksum")
        'intSIP2Port = tblInfo.Rows(0).Item("SIP2Port")
        'ConnectEnable = True

        ServerWinsock = New clsServerWinsock(6001)
    End Sub
    Private Sub DisposeAllSip2Patron()
        Dim i As Integer
        For i = 0 To Max_Client
            If Not objSip2Patron(i) Is Nothing Then
                objSip2Patron(i).Dispose(True)
                objSip2Patron(i) = Nothing
            End If
        Next
    End Sub

    Private Sub AddnewSip2Client(ByVal index As Integer)
        If objSip2Patron(index) Is Nothing Then
            objSip2Patron(index) = New clsSip2Patron
            objSip2Patron(index).Connectionstring = ConnectionString
            objSip2Patron(index).DBServer = DBServer
            Call objSip2Patron(index).Initialize()
        End If
    End Sub

    ' Method: GetXmlFile
    Public Function GetXmlFile(ByVal strFileNameXml As String) As DataTable
        ' Use function ConvertTable
        Dim strName As String = ""
        Dim dsResource As New DataSet
        Dim dt As DataTable = Nothing
        Try
            dsResource.ReadXml(strFileNameXml)
            If dsResource.Tables.Count > 0 Then
                dt = dsResource.Tables(0)
                dsResource.Tables.Clear()
            End If
        Catch ex As Exception
        Finally
        End Try
        Return dt
    End Function

    Private Sub WriteSip2Log(ByVal strMessage As String, Optional ByVal blnCreate As Boolean = False)
        Dim fsLog As StreamWriter
        If blnCreate Then
            fsLog = File.CreateText("C:\Sip2Server.log")
        Else
            fsLog = File.AppendText("C:\Sip2Server.log")
        End If
        fsLog.WriteLine(strMessage)
        fsLog.Close()
        fsLog = Nothing
    End Sub
End Module
