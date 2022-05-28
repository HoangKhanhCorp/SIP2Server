Imports System.ServiceProcess
Imports WinSocket
Imports System.IO
Imports System.Windows.Forms
Imports ProcessData
Imports LibolLogin
Imports System.Net
Imports System.Net.Sockets

Public Class Sip2Service
    Inherits System.ServiceProcess.ServiceBase

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call

    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New Sip2Service}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
        Me.ServiceName = "Sip2Service"
    End Sub

#End Region
    Private arrRequest(Max_Client) As String
    Private arrLastResponse(Max_Client) As String
    Private arrClient(Max_Client) As String
    Private strTimeZone As String
    Private strCurrencyUnit As String
    Private strChecksumEnabled As String
    Private intSIP2Port As Integer
    '  Private WithEvents ServerWinsock As clsServerWinsock
    'AppMain_OnMessage
    Private WithEvents ServerWinsock As clsServerWinsock
    Private clientSocket As Socket

    Private Const Max_Client As Integer = 10
    Private objSip2Patron(Max_Client) As clsSip2Patron

    Private ConnectionString As String
    Private ConnectEnable As Boolean
    Private DBServer As String
    Private Sub DisposeAllSip2Patron()
        Dim i As Integer
        For i = 0 To Max_Client
            If Not objSip2Patron(i) Is Nothing Then
                objSip2Patron(i).Dispose(True)
                objSip2Patron(i) = Nothing
            End If
        Next
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        Try
            Dim tblInfo As DataTable
            ' Dim objLg As New LibolLogin.clsLibolLogin
            'ConnectionString = objLg.GetConnectionString()
            'DBServer = objLg.DBServer
            tblInfo = GetXmlFile(Application.StartupPath & "\Sip2Server.xml")
            ConnectionString = tblInfo.Rows(0).Item("ConnectString")
            DBServer = "SQLSERVER"
            strTimeZone = tblInfo.Rows(0).Item("TimeZone")
            strCurrencyUnit = tblInfo.Rows(0).Item("Currency")
            strChecksumEnabled = tblInfo.Rows(0).Item("Checksum")
            intSIP2Port = tblInfo.Rows(0).Item("SIP2Port")
            ConnectEnable = True

            ServerWinsock = New clsServerWinsock(intSIP2Port)
            '      ServerWinsock = New AppMain()
            'AddnewSip2Client(0)
            'If Not objSip2Patron(0).CheckOpenConnection Then
            '    WriteSip2Log("Not connect")
            '    MsgBox("There's an error while connecting to database!", , "Database Connection Error !")
            '    ConnectEnable = False
            'End If
            'ServerWinsock = New clsServerWinsock(intSIP2Port, clientSocket)
            'Dim inti As Integer
            'For inti = 0 To Max_Client
            '    arrRequest(inti) = ""
            '    arrLastResponse(inti) = ""
            '    arrClient(inti) = ""
            'Next
            'If ConnectEnable Then
            '    WriteSip2Log("Service start at " & Now.ToString, True)
            'Else
            '    WriteSip2Log("Not connect1")
            '    WriteSip2Log("There's an error while connecting to database!", True)
            '    WriteSip2Log(ConnectionString)
            'End If
        Catch ex As Exception
            WriteSip2Log("EX:" & ex.Message)
        End Try
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
    'SocketChatClient
    '  Private Sub ServerWinsock_OnMessage(ByVal IncomingMessage As String, ByVal strIpClient As String) Handles ServerWinsock.OnMessage
    Private Sub ServerWinsock_OnMessage(ByVal IncomingMessage As String, ByVal strIpClient As String) Handles ServerWinsock.OnMessage
        If Not ConnectEnable Then
            Exit Sub
        End If
        Dim Buffer() As Byte
        Dim intIndex As Integer
        Dim blnHave As Boolean = False
        Dim strData As String = System.Text.Encoding.UTF8.GetString(IncomingMessage) ' IncomingMessage
        Dim strResponse As String = ""
        If strData.Length = 0 Then
            Exit Sub
        End If
        ' kiem tra thong tin tu client nao toi
        For intIndex = 0 To Max_Client
            If arrClient(intIndex) = strIpClient And strIpClient <> "" Then
                blnHave = True
                Exit For
            End If
            If arrClient(intIndex) = "" Then
                Exit For
            End If
        Next
        If Not blnHave And strIpClient <> "" Then
            arrClient(intIndex) = strIpClient
            AddnewSip2Client(intIndex)
        End If

        'If Left(strData, 7).ToLower = "connect" Then
        '    objSip2Patron(intIndex).UserID = 0
        '    WriteSip2Log("R-" & strIpClient & ": Connected")
        '    Exit Sub
        'End If
        'WriteSip2Log("R-" & strIpClient & ":" & Left(strData, InStr(strData, Chr(13)) - 1))

        'bat dau xu ly thong tin
        If InStr(strData, Chr(13)) > 0 Then
            WriteSip2Log("data:" & strData)
            arrRequest(intIndex) = arrRequest(intIndex) & Left(strData, InStr(strData, Chr(13)) - 1)
            strResponse = ProcessRequest(arrRequest(intIndex), intIndex)
            arrLastResponse(intIndex) = strResponse
            Buffer = System.Text.Encoding.ASCII.GetBytes(strResponse.ToCharArray)
            '   Dim objSocketServer As Socket = ServerWinsock.GetSocketServer()
            '  WriteSip2Log("services:" & strResponse)
            '
            sc.Send(Buffer)

            '    client.Sock.Send(byteDateLine, byteDateLine.Length, 0)

            arrRequest(intIndex) = ""
        Else
            arrRequest(intIndex) = arrRequest(intIndex) & strData
        End If
    End Sub

    Private Sub CloseAll()
        Call DisposeAllSip2Patron()
        ServerWinsock = Nothing
        'Me.Dispose(True)
        'End
    End Sub
    Private Function Checksum(ByVal strIn As String) As String
        Dim intDecChecksum, i As Integer
        Dim strHexChecksum As String
        intDecChecksum = 0
        For i = 1 To Len(strIn)
            intDecChecksum = intDecChecksum + Asc(Mid(strIn, i, 1))
        Next
        strHexChecksum = Hex((Not intDecChecksum) + 1)
        Checksum = strHexChecksum.PadLeft(4, "0")
    End Function

    Private Function VerifyChecksum(ByVal strIn As String) As Boolean
        Dim strHexChecksum, strIncludedChecksum As String
        strHexChecksum = Checksum(Left(strIn, Len(strIn) - 4))
        strIncludedChecksum = Right(strIn, 4)
        If strIncludedChecksum = strHexChecksum Then
            VerifyChecksum = True
        Else
            VerifyChecksum = False
        End If
    End Function

    Private Function ProcessRequest(ByVal req As String, ByVal idx As Integer) As String
        Dim strCommandId As String = String.Empty
        Dim strControl As String = String.Empty
        Dim strChecksum As String = String.Empty
        Dim strSequence As String = String.Empty
        Dim strOutput As String = String.Empty

        strChecksum = strChecksumEnabled
        If strChecksum.ToLower = "yes" Then
            strControl = Right(req, 9)
            If InStr(strControl, "AY") = 1 And InStr(strControl, "AZ") = 4 Then
                If Not VerifyChecksum(req) Then
                    '  ProcessRequest = "96AZ" & Checksum("96AZ")
                    ProcessRequest = "98YYYYYN3000052020022407001156032.00AO|AM|BXYYYYYYYYYYYYYYYY|AN|AF|AG"
                    Exit Function
                Else
                    strSequence = Mid(req, Len(req) - 6, 1)
                    req = Left(req, Len(req) - 9)
                End If
            Else
                strChecksum = "no"
            End If
        End If
        If req = "" Then
            Return ""
            Exit Function
        End If
        strCommandId = Left(req, 2)
        req = Right(req, Len(req) - 2)
        If strCommandId <> 93 And objSip2Patron(idx).UserID = 0 Then
            strOutput = "940"
        Else
            Select Case strCommandId
                ' Patron status request
                Case "23"
                    strOutput = objSip2Patron(idx).PatronStatusRequest(req, True)
                    ' Checkout
                Case "11"
                    strOutput = objSip2Patron(idx).CheckOut(req)
                    ' Checkin
                Case "09"
                    strOutput = objSip2Patron(idx).CheckIn(req)
                    ' Block patron
                Case "01"
                    strOutput = objSip2Patron(idx).BlockPatron(req)
                    ' SC Status
                Case "99"
                    strOutput = objSip2Patron(idx).SCStatus(req)
                    ' Request ACS resend
                Case "97"
                    strOutput = arrLastResponse(idx)
                    ' Login
                Case "93"
                    strOutput = objSip2Patron(idx).Login(req)
                    ' Patron Information
                Case "63"
                    strOutput = objSip2Patron(idx).PatronInformation(req)
                    ' End Patron Session
                Case "35"
                    strOutput = objSip2Patron(idx).EndPatronSession(req)
                    ' Fee Paid
                Case "37"
                    strOutput = objSip2Patron(idx).FeePaid(req)
                    ' Item Information
                Case "17"
                    strOutput = objSip2Patron(idx).ItemInformation(req)
                    ' Item Status Update
                Case "19"
                    strOutput = objSip2Patron(idx).ItemStatusUpdate(req)
                    ' Patron enable
                Case "25"
                    strOutput = objSip2Patron(idx).PatronEnable(req)
                    ' Hold
                Case "15"
                    strOutput = objSip2Patron(idx).Hold(req)
                    ' Renew
                Case "29"
                    strOutput = objSip2Patron(idx).Renew(req)
                    ' Renew all
                Case "65"
                    strOutput = objSip2Patron(idx).RenewAll(req)
            End Select
        End If
        If strChecksum = "yes" Then
            strOutput = strOutput & "AY" & strSequence & "AZ"
            strOutput = strOutput & Checksum(strOutput)
        End If
        ProcessRequest = strOutput
    End Function

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.

        Call CloseAll()
    End Sub

End Class
