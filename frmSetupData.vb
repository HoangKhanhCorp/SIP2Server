Imports ProcessData
Public Class frmSetupData
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdSqlserver As System.Windows.Forms.RadioButton
    Friend WithEvents rdOracle As System.Windows.Forms.RadioButton
    Friend WithEvents grSQLServer As System.Windows.Forms.GroupBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtMatkhau As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtTen As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCSDL As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtMaychu As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents grpOracle As System.Windows.Forms.GroupBox
    Friend WithEvents txtOraPass As System.Windows.Forms.TextBox
    Friend WithEvents txtOraUsr As System.Windows.Forms.TextBox
    Friend WithEvents txtOraDS As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rdOracle = New System.Windows.Forms.RadioButton
        Me.rdSqlserver = New System.Windows.Forms.RadioButton
        Me.grSQLServer = New System.Windows.Forms.GroupBox
        Me.txtMatkhau = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtTen = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtCSDL = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtMaychu = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnConnect = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.grpOracle = New System.Windows.Forms.GroupBox
        Me.txtOraPass = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtOraUsr = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtOraDS = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.grSQLServer.SuspendLayout()
        Me.grpOracle.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdOracle)
        Me.GroupBox1.Controls.Add(Me.rdSqlserver)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(112, 128)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Database"
        '
        'rdOracle
        '
        Me.rdOracle.Location = New System.Drawing.Point(8, 76)
        Me.rdOracle.Name = "rdOracle"
        Me.rdOracle.Size = New System.Drawing.Size(88, 24)
        Me.rdOracle.TabIndex = 1
        Me.rdOracle.Text = "Oracle"
        '
        'rdSqlserver
        '
        Me.rdSqlserver.Checked = True
        Me.rdSqlserver.Location = New System.Drawing.Point(8, 48)
        Me.rdSqlserver.Name = "rdSqlserver"
        Me.rdSqlserver.Size = New System.Drawing.Size(88, 24)
        Me.rdSqlserver.TabIndex = 0
        Me.rdSqlserver.TabStop = True
        Me.rdSqlserver.Text = "SQL Server"
        '
        'grSQLServer
        '
        Me.grSQLServer.Controls.Add(Me.txtMatkhau)
        Me.grSQLServer.Controls.Add(Me.Label4)
        Me.grSQLServer.Controls.Add(Me.txtTen)
        Me.grSQLServer.Controls.Add(Me.Label3)
        Me.grSQLServer.Controls.Add(Me.txtCSDL)
        Me.grSQLServer.Controls.Add(Me.Label2)
        Me.grSQLServer.Controls.Add(Me.txtMaychu)
        Me.grSQLServer.Controls.Add(Me.Label1)
        Me.grSQLServer.Location = New System.Drawing.Point(128, 8)
        Me.grSQLServer.Name = "grSQLServer"
        Me.grSQLServer.Size = New System.Drawing.Size(296, 128)
        Me.grSQLServer.TabIndex = 1
        Me.grSQLServer.TabStop = False
        Me.grSQLServer.Text = "SQL Server"
        Me.grSQLServer.Visible = False
        '
        'txtMatkhau
        '
        Me.txtMatkhau.Location = New System.Drawing.Point(88, 96)
        Me.txtMatkhau.Name = "txtMatkhau"
        Me.txtMatkhau.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtMatkhau.Size = New System.Drawing.Size(200, 20)
        Me.txtMatkhau.TabIndex = 23
        Me.txtMatkhau.Text = "libol50"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Password:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtTen
        '
        Me.txtTen.Location = New System.Drawing.Point(88, 72)
        Me.txtTen.Name = "txtTen"
        Me.txtTen.Size = New System.Drawing.Size(200, 20)
        Me.txtTen.TabIndex = 21
        Me.txtTen.Text = "libol50"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Username:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCSDL
        '
        Me.txtCSDL.Location = New System.Drawing.Point(88, 48)
        Me.txtCSDL.Name = "txtCSDL"
        Me.txtCSDL.Size = New System.Drawing.Size(200, 20)
        Me.txtCSDL.TabIndex = 19
        Me.txtCSDL.Text = "Libol601m"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Data Source:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMaychu
        '
        Me.txtMaychu.Location = New System.Drawing.Point(88, 24)
        Me.txtMaychu.Name = "txtMaychu"
        Me.txtMaychu.Size = New System.Drawing.Size(200, 20)
        Me.txtMaychu.TabIndex = 17
        Me.txtMaychu.Text = "192.168.50.199"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "IP Host:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(16, 144)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(120, 24)
        Me.btnOk.TabIndex = 2
        Me.btnOk.Text = "&Ok"
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(152, 144)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(120, 24)
        Me.btnConnect.TabIndex = 3
        Me.btnConnect.Text = "&Test Connection"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(296, 144)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(120, 24)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        '
        'grpOracle
        '
        Me.grpOracle.Controls.Add(Me.txtOraPass)
        Me.grpOracle.Controls.Add(Me.Label5)
        Me.grpOracle.Controls.Add(Me.txtOraUsr)
        Me.grpOracle.Controls.Add(Me.Label6)
        Me.grpOracle.Controls.Add(Me.txtOraDS)
        Me.grpOracle.Controls.Add(Me.Label7)
        Me.grpOracle.Location = New System.Drawing.Point(48, 192)
        Me.grpOracle.Name = "grpOracle"
        Me.grpOracle.Size = New System.Drawing.Size(296, 128)
        Me.grpOracle.TabIndex = 5
        Me.grpOracle.TabStop = False
        Me.grpOracle.Text = "Oracle"
        Me.grpOracle.Visible = False
        '
        'txtOraPass
        '
        Me.txtOraPass.Location = New System.Drawing.Point(84, 82)
        Me.txtOraPass.Name = "txtOraPass"
        Me.txtOraPass.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtOraPass.Size = New System.Drawing.Size(200, 20)
        Me.txtOraPass.TabIndex = 15
        Me.txtOraPass.Text = "libol60"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(4, 82)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 16)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Password:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOraUsr
        '
        Me.txtOraUsr.Location = New System.Drawing.Point(84, 58)
        Me.txtOraUsr.Name = "txtOraUsr"
        Me.txtOraUsr.Size = New System.Drawing.Size(200, 20)
        Me.txtOraUsr.TabIndex = 13
        Me.txtOraUsr.Text = "libol601m"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(12, 58)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Username:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOraDS
        '
        Me.txtOraDS.Location = New System.Drawing.Point(84, 34)
        Me.txtOraDS.Name = "txtOraDS"
        Me.txtOraDS.Size = New System.Drawing.Size(200, 20)
        Me.txtOraDS.TabIndex = 11
        Me.txtOraDS.Text = "Libol"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(12, 34)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 16)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Data Source:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmSetupData
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.PeachPuff
        Me.ClientSize = New System.Drawing.Size(432, 181)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnConnect)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.grSQLServer)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.grpOracle)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSetupData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Database Connection"
        Me.GroupBox1.ResumeLayout(False)
        Me.grSQLServer.ResumeLayout(False)
        Me.grpOracle.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmSetupData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim locPoint As Point
        locPoint.X = 128
        locPoint.Y = 8
        grpOracle.Location = locPoint
        If DBServer = "ORACLE" Then
            rdOracle.Checked = True
            rdSqlserver.Checked = False
            txtOraDS.Text = DataSource
            txtOraUsr.Text = UserName
            txtOraPass.Text = Password
            grpOracle.Visible = True
            grSQLServer.Visible = False
        Else
            rdOracle.Checked = False
            rdSqlserver.Checked = True
            txtMaychu.Text = IpHost
            txtCSDL.Text = DataSource
            txtTen.Text = UserName
            txtMatkhau.Text = Password
            grpOracle.Visible = False
            grSQLServer.Visible = True
        End If
    End Sub

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Dim strConnectionString As String
        Dim objBTestConnect As New clsSip2Patron
        If rdSqlserver.Checked Then
            objBTestConnect.DBServer = "SQLSERVER"
            objBTestConnect.Connectionstring = "Data Source=" & txtMaychu.Text & ";Initial Catalog=" & txtCSDL.Text & ";UID=" & txtTen.Text & ";PWD=" & txtMatkhau.Text & ";"
        Else
            objBTestConnect.DBServer = "ORACLE"
            objBTestConnect.Connectionstring = "User ID=" & txtOraUsr.Text & ";Password=" & txtOraPass.Text & ";Data Source=" & txtOraDS.Text
        End If
        Call objBTestConnect.Initialize()
        If Not objBTestConnect.CheckOpenConnection Then
            MsgBox("Could not connected to database!", , "Database Connection Error !")
        Else
            MsgBox("Connected successfull!", , "Connection Ok!")
        End If
        Call objBTestConnect.Dispose(True)
        objBTestConnect = Nothing
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        If rdSqlserver.Checked Then
            DBServer = "SQLSERVER"
            IpHost = txtMaychu.Text
            DataSource = txtCSDL.Text
            UserName = txtTen.Text
            Password = txtMatkhau.Text
            ConnectionString = "Data Source=" & IpHost & ";Initial Catalog=" & DataSource & ";UID=" & UserName & ";PWD=" & Password & ";"
        Else
            DBServer = "ORACLE"
            DataSource = txtOraDS.Text
            UserName = txtOraUsr.Text
            Password = txtOraPass.Text
            ConnectionString = "User ID=" & UserName & ";Password=" & Password & ";Data Source=" & DataSource
        End If
        Call DisposeAllSip2Patron()
        objSip2Patron(0) = New clsSip2Patron
        objSip2Patron(0).Connectionstring = ConnectionString
        objSip2Patron(0).DBServer = DBServer
        Call objSip2Patron(0).Initialize()
        blnConnectEnable = True
        If Not objSip2Patron(0).CheckOpenConnection Then
            MsgBox("There's an error while connecting to database!", , "Database Connection Error !")
            blnConnectEnable = False
        End If
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub rdSqlserver_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdSqlserver.CheckedChanged
        grpOracle.Visible = False
        grSQLServer.Visible = True
    End Sub

    Private Sub rdOracle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOracle.CheckedChanged
        grpOracle.Visible = True
        grSQLServer.Visible = False
    End Sub
End Class
