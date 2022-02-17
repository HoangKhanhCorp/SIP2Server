Imports System.ServiceProcess
Imports WinSocket
Imports System.IO
Imports System.Windows.Forms
Imports ProcessData
Imports LibolLogin
Imports System.Net
Imports System.Net.Sockets
Public Class FormTest
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
    Private NameDonVi As String
    Private strIPAddress As String
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
    Private Sub Showinfo(str As String, Optional lrequest As Boolean = 1, Optional lrReponse As Boolean = 1)
        lbl.Text = str
        'If lrequest = 1 OrElse lrReponse = 1 Then
        '    txtLog.Text = txtLog.Text & vbCrLf & str
        'End If
        txtLog.Text = txtLog.Text & vbCrLf & str
    End Sub
    Private Function GetIP() As IPAddress
        Dim strHostName As String = Dns.GetHostName
        Dim IpHostEntry As IPHostEntry = Dns.GetHostByName(strHostName)
        Dim IpAddress As IPAddress = IpHostEntry.AddressList(1)
        Return IpAddress
    End Function
    Private Sub FormTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim tblInfo As DataTable
        'Dim objLg As New LibolLogin.clsLibolLogin
        'ConnectionString = objLg.GetConnectionString()
        'DBServer = objLg.DBServer
        ' ConnectionString = "SERVER=.\SQL2017;UID=sa;PWD=123;DATABASE=LibolTDM"
        DBServer = "SQLSERVER"
        tblInfo = GetXmlFile(Application.StartupPath & "\Sip2Server.xml")
        strTimeZone = "    " 'tblInfo.Rows(0).Item("TimeZone") '"" '

        strCurrencyUnit = tblInfo.Rows(0).Item("Currency")
        strIPAddress = tblInfo.Rows(0).Item("IP")
        strChecksumEnabled = tblInfo.Rows(0).Item("Checksum") '"yes" ' 
        intSIP2Port = tblInfo.Rows(0).Item("SIP2Port")
        ConnectionString = tblInfo.Rows(0).Item("ConnectString")
        NameDonVi = tblInfo.Rows(0).Item("LibName")
        ConnectEnable = True

        Call Showinfo(ConnectionString)
        If strIPAddress = "" Then
            ServerWinsock = New clsServerWinsock(intSIP2Port)
            Call Showinfo("Listening: " & GetIP.ToString())
        Else
            ServerWinsock = New clsServerWinsock(intSIP2Port, strIPAddress)
            Call Showinfo("Listening: " & strIPAddress)
        End If







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
    End Sub

    Private Sub AddnewSip2Client(ByVal index As Integer)
        If objSip2Patron(index) Is Nothing Then
            objSip2Patron(index) = New clsSip2Patron
            objSip2Patron(index).Connectionstring = ConnectionString
            objSip2Patron(index).DBServer = DBServer
            objSip2Patron(index).UserID = 1
            objSip2Patron(index).TimeZone = strTimeZone
            objSip2Patron(index).LibName = NameDonVi
            Call objSip2Patron(index).Initialize()
        End If
    End Sub

    ' Method: GetXmlFile
    Public Function GetXmlFile(ByVal strFileNameXml As String) As DataTable
        ' Use function ConvertTable
        Dim strName As String = ""
        Dim dsResource As New DataSet
        Try
            dsResource.ReadXml(strFileNameXml)
            If dsResource.Tables.Count > 0 Then
                GetXmlFile = dsResource.Tables(0)
                dsResource.Tables.Clear()
            End If
        Catch ex As Exception
        Finally
        End Try
    End Function

    Private Sub WriteSip2Log(ByVal strMessage As String, Optional ByVal blnCreate As Boolean = False)
        Dim fsLog As StreamWriter
        Try
            If blnCreate Then
                fsLog = File.CreateText(Application.StartupPath + "\Sip2Server.log")
            Else
                fsLog = File.AppendText(Application.StartupPath + "\Sip2Server.log")
            End If
        Catch ex As Exception

        Finally
            fsLog.WriteLine(strMessage)
            fsLog.Close()
            fsLog = Nothing
        End Try


    End Sub

    Private Sub ServerWinsock_OnMessage(ByVal IncomingMessage As String, ByVal strIpClient As String) Handles ServerWinsock.OnMessage
        If Not ConnectEnable Then
            Exit Sub
        End If
        Dim Buffer() As Byte
        Dim intIndex As Integer
        Dim blnHave As Boolean = False
        Dim strData As String = IncomingMessage
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
            arrRequest(intIndex) = arrRequest(intIndex) & strData.Substring(0, InStr(strData, Chr(13)) - 1)

            WriteSip2Log("Request: " & DateTime.Now.ToLongTimeString() & " -Content: " & arrRequest(intIndex))
            strResponse = ProcessRequest(arrRequest(intIndex), intIndex)
            arrLastResponse(intIndex) = strResponse
            Buffer = System.Text.Encoding.ASCII.GetBytes(strResponse.ToCharArray)
            Dim objSocketServer As Socket = ServerWinsock.GetSocketServer()
            objSocketServer.Send(Buffer)


            WriteSip2Log("services: " & DateTime.Now.ToLongTimeString() & " -Content: " & strResponse)

            arrRequest(intIndex) = ""
        Else
            arrRequest(intIndex) = arrRequest(intIndex) & strData
        End If
    End Sub

    Private Sub CloseAll()
        Call DisposeAllSip2Patron()
        ServerWinsock.Dispose(True)
        ServerWinsock = Nothing
        'Me.Dispose(True)
        'End
    End Sub
    Private Function Checksum(ByVal strIn As String) As String
        Dim intDecChecksum, i As Integer
        Dim strHexChecksum As String
        intDecChecksum = 0
        For i = 1 To Len(strIn)
            intDecChecksum = intDecChecksum + Asc(Mid(strIn, i, 1)) '+ strIn.Substring(i, 1) '
        Next
        strHexChecksum = Hex((Not intDecChecksum) + 1)
        '  Checksum = strHexChecksum.PadLeft(4, "0")
        Checksum = strHexChecksum.Substring(strHexChecksum.Length - 4, 4)

    End Function
    Public Function AppendChecksum(ByVal strMsg As String) As String

        Dim intCtr As Integer
        Dim chrArray As Char()
        Dim intAscSum As Integer
        Dim blnCarryBit As Boolean
        Dim strBinVal As String
        Dim strInvBinVal As String
        Dim strNewBinVal As String

        'Transfer SIP message to a a character array. Loop through each character of the array, 
        'converting the character to an ASCII value and adding the value to a running total. 

        intAscSum = 0
        chrArray = strMsg.ToCharArray

        For intCtr = 0 To chrArray.Length - 1
            intAscSum = intAscSum + Asc(chrArray(intCtr))
        Next


        'Next, convert ASCII sum to a binary digit by: 
        ' 1) taking the remainder of the ASCII sum divided by 2 
        ' 2) Performing integer division by 2 on the sum (using "\" instead of "/") 
        ' 3) Repeat until sum reaches 0 
        ' 4) Pad to 16 digits with leading zeroes 

        Do
            strBinVal = CStr(intAscSum Mod 2) & strBinVal
            intAscSum = intAscSum \ 2
        Loop Until intAscSum = 0

        strBinVal = strBinVal.PadLeft(16, "0")

        'Next, invert all bits in binary number. 
        chrArray = strBinVal.ToCharArray
        strInvBinVal = ""

        For intCtr = 0 To chrArray.Length - 1
            If chrArray(intCtr) = "0" Then
                strInvBinVal = strInvBinVal & "1"
            Else
                strInvBinVal = strInvBinVal & "0"
            End If
        Next


        'Next, add 1 to the inverted binary digit. Loop from least significant digit (rightmost) to most (leftmost); 
        'if digit is 1, flip to 0 and retain carry bit to next significant digit. 

        blnCarryBit = True
        chrArray = strInvBinVal.ToCharArray

        For intCtr = chrArray.Length - 1 To 0 Step -1
            If blnCarryBit = True Then
                If chrArray(intCtr) = "0" Then
                    chrArray(intCtr) = "1"
                    blnCarryBit = False
                Else
                    chrArray(intCtr) = "0"
                    blnCarryBit = True
                End If
            End If

            strNewBinVal = chrArray(intCtr) & strNewBinVal
        Next

        'Finally, convert binary digit to hex value, append to original SIP message. 

        AppendChecksum = strMsg & Hex(Convert.ToInt64(strNewBinVal, 2))

    End Function
    'Private Function Checksum(ByVal strIn As String) As String
    '    Dim intDecChecksum, i As Integer
    '    Dim strHexChecksum As String
    '    intDecChecksum = 0
    '    For i = 1 To Len(strIn)
    '        intDecChecksum = intDecChecksum + Asc(Mid(strIn, i, 1)) '+ strIn.Substring(i, 1) '
    '    Next
    '    strHexChecksum = Hex((Not intDecChecksum) + 1)
    '    Checksum = strHexChecksum.PadLeft(4, "0")
    'End Function
    Private Function VerifyChecksum(ByVal strIn As String) As Boolean
        Dim strHexChecksum, strIncludedChecksum As String
        strHexChecksum = Checksum(strIn.Substring(0, Len(strIn) - 4))
        ' strHexChecksum = Checksum(Right(strIn, 4))
        strIncludedChecksum = strIn.Substring(strIn.Length - 4, 4)
        If strIncludedChecksum = strHexChecksum Then
            VerifyChecksum = True
        Else
            VerifyChecksum = False
        End If
    End Function
    Private iSeqNumber As Int16 = 0
    Private Function ProcessRequest(ByVal req As String, ByVal idx As Integer) As String
        Dim strCommandId As String = String.Empty
        Dim strControl As String = String.Empty
        Dim strChecksum As String = String.Empty
        Dim strSequence As String = String.Empty
        Dim strOutput As String = String.Empty

        strChecksum = strChecksumEnabled
        If strChecksum.ToLower = "yes" Then
            If req.Length > 9 Then
                'Check lai nha
                strControl = req.Substring(req.Length - 9, 9)
            Else
                strControl = req
            End If
            'Check lai nha
            '  If InStr(strControl, "AY") = 1 And InStr(strControl, "AZ") = 4 Then
            'InStr(strControl, "AC") = 1 And
            If InStr(strControl, "AY") = 1 And InStr(strControl, "AZ") = 4 Then
                If Not VerifyChecksum(req) Then
                    ProcessRequest = "96AZ" & Checksum("96AZ")
                    Exit Function
                Else
                    strSequence = Mid(req, Len(req) - 9, 9)
                    req = req.Substring(0, Len(req) - 9)
                End If
            Else
                strChecksum = "no"
            End If
        End If
        If req = "" Then
            Return ""
            Exit Function
        End If
        strCommandId = req.Substring(0, 2)
        'req = Right(req, Len(req) - 2)
        req = req.Substring(2, Len(req) - 2)
        If strCommandId <> "93" And objSip2Patron(idx).UserID = 0 Then
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
            'strOutput = strOutput & "|AY" & strSequence & "AZ"
            'strOutput = strOutput & Checksum(strOutput) & vbCrLf
            '  strOutput = strOutput & AppendChecksum(strOutput) & vbCrLf

        End If
        iSeqNumber = iSeqNumber + 1
        If iSeqNumber > 9 Then
            iSeqNumber = 0
        End If
        strOutput = AppendChecksum(strOutput & "|AY" & iSeqNumber & "AZ") & vbCrLf
        ProcessRequest = strOutput
    End Function

    Private Sub FormTest_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.CloseAll()
    End Sub
End Class