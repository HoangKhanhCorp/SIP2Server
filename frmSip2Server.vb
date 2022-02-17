Imports System.IO
Imports WinSocket
Imports ProcessData
Public Class frmSip2Server
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ntIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents ctMenu As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents lstResult As System.Windows.Forms.ListBox
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents mmSystem As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSip2Server))
        Me.ntIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ctMenu = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.lstResult = New System.Windows.Forms.ListBox
        Me.mmSystem = New System.Windows.Forms.MainMenu
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'ntIcon
        '
        Me.ntIcon.ContextMenu = Me.ctMenu
        Me.ntIcon.Icon = CType(resources.GetObject("ntIcon.Icon"), System.Drawing.Icon)
        Me.ntIcon.Text = "Libol Sip2 Server is running..."
        '
        'ctMenu
        '
        Me.ctMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem5, Me.MenuItem12, Me.MenuItem2, Me.MenuItem3, Me.MenuItem10})
        '
        'MenuItem1
        '
        Me.MenuItem1.DefaultItem = True
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Sho&w Sip2Server"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "&Stop Sip2 Server"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 2
        Me.MenuItem12.Text = "Setup &Database"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 4
        Me.MenuItem3.Text = "&About"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 5
        Me.MenuItem10.Text = "&Exit"
        '
        'lstResult
        '
        Me.lstResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstResult.BackColor = System.Drawing.SystemColors.InfoText
        Me.lstResult.ForeColor = System.Drawing.Color.GreenYellow
        Me.lstResult.Location = New System.Drawing.Point(0, 0)
        Me.lstResult.Name = "lstResult"
        Me.lstResult.Size = New System.Drawing.Size(768, 342)
        Me.lstResult.TabIndex = 8
        '
        'mmSystem
        '
        Me.mmSystem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4})
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem11, Me.MenuItem8, Me.MenuItem7, Me.MenuItem9})
        Me.MenuItem4.Text = "&System"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Stop Sip2 Server"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 1
        Me.MenuItem11.Text = "Setup &Database"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "-"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 3
        Me.MenuItem7.Text = "&About"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 4
        Me.MenuItem9.Text = "&Exit"
        '
        'frmSip2Server
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(768, 345)
        Me.Controls.Add(Me.lstResult)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.mmSystem
        Me.Name = "frmSip2Server"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Libol Sip2 Server"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private arrRequest(Max_Client) As String
    Private arrLastResponse(Max_Client) As String
    Private arrClient(Max_Client) As String
    Private strTimeZone As String
    Private strCurrencyUnit As String
    Private strChecksumEnabled As String
    Private intSIP2Port As Integer
    Private WithEvents ServerWinsock As clsServerWinsock
    Private ClientWinsock As clsClientWinsock
    Private blnStartSip As Boolean
    Private Sub frmSip2Server_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim tblInfo As DataTable
        tblInfo = GetXmlFile("Sip2Server.xml")
        SetConnectionString(tblInfo.Rows(0).Item("FileNameDB"))
        strTimeZone = tblInfo.Rows(0).Item("TimeZone")
        strCurrencyUnit = tblInfo.Rows(0).Item("Currency")
        strChecksumEnabled = tblInfo.Rows(0).Item("Checksum")
        intSIP2Port = tblInfo.Rows(0).Item("SIP2Port")
        blnConnectEnable = True
        AddnewSip2Client(0)
        If Not objSip2Patron(0).CheckOpenConnection Then
            MsgBox("There's an error while connecting to database!", , "Database Connection Error !")
            blnConnectEnable = False
        End If
        If blnConnectEnable Then
            'start server listener
            ServerWinsock = New clsServerWinsock(intSIP2Port)
            blnStartSip = True
        Else
            mmSystem.MenuItems(0).MenuItems(0).Text = "&Start Sip2 Server"
            ctMenu.MenuItems(1).Text = "&Start Sip2 Server"
            ntIcon.Text = "Libol Sip2 Server is stopped..."
            blnStartSip = False
        End If
        Dim inti As Integer
        For inti = 0 To Max_Client
            arrRequest(inti) = ""
            arrLastResponse(inti) = ""
            arrClient(inti) = ""
        Next
        ntIcon.Visible = True
        Me.Visible = False
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

    ' Method: ResetSession
    Private Sub SetConnectionString(ByVal strPathLogin As String)
        Dim tblTest As DataTable
        Dim objXe As XCrypt.XCryptEngine
        Dim strEncript As String = ""
        Dim inti As Integer
        'Put user code to initialize the page here
        tblTest = GetXmlFile(strPathLogin)
        If Not tblTest Is Nothing AndAlso tblTest.Rows.Count > 0 Then
            For inti = 0 To tblTest.Rows.Count - 1
                If tblTest.Rows(inti).Item("Run") = 1 Then
                    If tblTest.Rows(inti).Item("DBServer") = "SQLSERVER" Then
                        DBServer = "SQLSERVER"
                        UserName = tblTest.Rows(inti).Item("UserName")
                        DataSource = tblTest.Rows(inti).Item("DataSource")
                        IpHost = tblTest.Rows(inti).Item("ServerIP")
                        objXe = New XCrypt.XCryptEngine(XCrypt.XCryptEngine.AlgorithmType.TripleDES)
                        Password = objXe.Decrypt(tblTest.Rows(inti).Item("PassWord"))
                        ConnectionString = "Data Source=" & IpHost & ";Initial Catalog=" & DataSource & ";UID=" & UserName & ";PWD=" & Password & ";"
                    Else
                        DBServer = "ORACLE"
                        UserName = tblTest.Rows(inti).Item("UserName")
                        DataSource = tblTest.Rows(inti).Item("DataSource")
                        IpHost = tblTest.Rows(inti).Item("ServerIP")
                        objXe = New XCrypt.XCryptEngine(XCrypt.XCryptEngine.AlgorithmType.TripleDES)
                        Password = objXe.Decrypt(tblTest.Rows(inti).Item("PassWord"))
                        ConnectionString = "User ID=" & UserName & ";Password=" & Password & ";Data Source=" & DataSource
                    End If
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Function RightStr(ByVal strInput As String, ByVal intLength As Integer) As String
        Return strInput.Substring(strInput.Length - intLength)
    End Function
    Private Function LeftStr(ByVal strInput As String, ByVal intLength As Integer) As String
        Return strInput.Substring(0, intLength)
    End Function

    Private Sub ServerWinsock_OnMessage(ByVal IncomingMessage As String, ByVal strIpClient As String) Handles ServerWinsock.OnMessage
        If Not blnConnectEnable Then
            Exit Sub
        End If
        Dim intIndex As Integer
        Dim blnHave As Boolean = False
        Dim strData As String = IncomingMessage
        Dim strResponse As String = ""
        intIndex = Len(Trim(strData))
        If intIndex < 2 Then
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

        If LeftStr(strData, 7).ToLower = "connect" Then
            ClientWinsock = New clsClientWinsock(strIpClient & ":" & CStr(intSIP2Port), "Sip2Server Connected Ok !")
            lstResult.Items.Add("R-" & strIpClient & ":" & strData)
            objSip2Patron(intIndex).UserID = 0
            Application.DoEvents()
            Exit Sub
        End If
        lstResult.Items.Add("R-" & strIpClient & ":" & LeftStr(strData, InStr(strData, Chr(13)) - 1))
        lstResult.SelectedIndex = lstResult.Items.Count - 1
        Application.DoEvents()
        'bat dau xu ly thong tin
        If InStr(strData, Chr(13)) > 0 Then
            arrRequest(intIndex) = arrRequest(intIndex) & LeftStr(strData, InStr(strData, Chr(13)) - 1)
            strResponse = ProcessRequest(arrRequest(intIndex), intIndex)
            arrLastResponse(intIndex) = strResponse
            ClientWinsock = New clsClientWinsock(strIpClient & ":" & CStr(intSIP2Port), strResponse)
            arrRequest(intIndex) = ""
        Else
            arrRequest(intIndex) = arrRequest(intIndex) & strData
        End If
    End Sub

    Private Sub CloseAll()
        Call DisposeAllSip2Patron()
        ServerWinsock = Nothing
        ntIcon.Visible = False
        ntIcon.Dispose()
        ntIcon = Nothing
        Me.Dispose(True)
        End
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
        strHexChecksum = Checksum(LeftStr(strIn, Len(strIn) - 4))
        strIncludedChecksum = RightStr(strIn, 4)
        If strIncludedChecksum = strHexChecksum Then
            VerifyChecksum = True
        Else
            VerifyChecksum = False
        End If
    End Function

    Private Function ProcessRequest(ByVal req As String, ByVal idx As Integer) As String
        Dim strCommandId, strControl, strChecksum, strSequence, strOutput As String
        strChecksum = strChecksumEnabled
        If strChecksum.ToLower = "yes" Then
            strControl = RightStr(req, 9)
            If InStr(strControl, "AY") = 1 And InStr(strControl, "AZ") = 4 Then
                If Not VerifyChecksum(req) Then
                    ProcessRequest = "96AZ" & Checksum("96AZ")
                    Exit Function
                Else
                    strSequence = Mid(req, Len(req) - 6, 1)
                    req = LeftStr(req, Len(req) - 9)
                End If
            Else
                strChecksum = "no"
            End If
        End If
        If req = "" Then
            Return ""
            Exit Function
        End If
        strCommandId = LeftStr(req, 2)
        req = RightStr(req, Len(req) - 2)
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

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim objAbout As New frmAbout
        objAbout.ShowDialog()
        objAbout = Nothing
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Call Me.StartSip2Server()
    End Sub

    Private Sub frmSip2Server_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Call CloseAll()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Me.Visible = True
        ntIcon.Visible = False
        Me.WindowState = FormWindowState.Normal
    End Sub


    Private Sub frmSip2Server_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Visible = False
            ntIcon.Visible = True
        End If
    End Sub

    Private Sub ntIcon_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntIcon.DoubleClick
        Me.Visible = True
        ntIcon.Visible = False
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Call StartSip2Server()
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Dim objAbout As New frmAbout
        objAbout.ShowDialog()
        objAbout = Nothing
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Call CloseAll()
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Call CloseAll()
    End Sub
    Private Sub StartSip2Server()
        If blnConnectEnable Then
            If blnStartSip Then
                mmSystem.MenuItems(0).MenuItems(0).Text = "&Start Sip2 Server"
                ctMenu.MenuItems(1).Text = "&Start Sip2 Server"
                ntIcon.Text = "Libol Sip2 Server is stopped..."
                ServerWinsock.Dispose(True)
                ServerWinsock = Nothing
                blnStartSip = False
            Else
                mmSystem.MenuItems(0).MenuItems(0).Text = "&Stop Sip2 Server"
                ctMenu.MenuItems(1).Text = "&Stop Sip2 Server"
                ntIcon.Text = "Libol Sip2 Server is running..."
                ServerWinsock = New clsServerWinsock(intSIP2Port)
                blnStartSip = True
            End If
        Else
            MsgBox("Could not start service ! There's an error while connecting to database!", , "Database Connection Error !")
        End If
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        If Not blnStartSip Then
            Dim frmSetupDB As New frmSetupData
            frmSetupDB.ShowDialog()
            frmSetupDB.Dispose()
            frmSetupDB = Nothing
        Else
            MsgBox("Could not opened while Sip2Server service is running!", , "Could not open!")
        End If
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        If Not blnStartSip Then
            Dim frmSetupDB As New frmSetupData
            frmSetupDB.ShowDialog()
            frmSetupDB.Dispose()
            frmSetupDB = Nothing
        Else
            MsgBox("Could not opened while Sip2Server service is running!", , "Could not open!")
        End If
    End Sub
End Class

