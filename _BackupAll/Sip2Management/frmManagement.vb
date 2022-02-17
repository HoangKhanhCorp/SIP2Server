Imports System.ServiceProcess
Imports System.IO
Public Class frmManagement
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
    Friend WithEvents ntiSip2 As System.Windows.Forms.NotifyIcon
    Friend WithEvents mmSip2 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents cmSip2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents tmSip2 As System.Windows.Forms.Timer
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents sbSip2 As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmManagement))
        Me.ntiSip2 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.cmSip2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.mmSip2 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.txtResult = New System.Windows.Forms.TextBox
        Me.tmSip2 = New System.Windows.Forms.Timer(Me.components)
        Me.sbSip2 = New System.Windows.Forms.StatusBar
        Me.SuspendLayout()
        '
        'ntiSip2
        '
        Me.ntiSip2.ContextMenu = Me.cmSip2
        Me.ntiSip2.Icon = CType(resources.GetObject("ntiSip2.Icon"), System.Drawing.Icon)
        Me.ntiSip2.Text = "Sip2Server service is running..."
        '
        'cmSip2
        '
        Me.cmSip2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem12, Me.MenuItem7, Me.MenuItem8, Me.MenuItem9, Me.MenuItem10, Me.MenuItem11})
        '
        'MenuItem12
        '
        Me.MenuItem12.DefaultItem = True
        Me.MenuItem12.Index = 0
        Me.MenuItem12.Text = "S&how Management"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "&Start Sip2 Server"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "S&top Sip2 Server"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 3
        Me.MenuItem9.Text = "-"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 4
        Me.MenuItem10.Text = "&About"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 5
        Me.MenuItem11.Text = "&Exit"
        '
        'mmSip2
        '
        Me.mmSip2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem5, Me.MenuItem6})
        Me.MenuItem1.Text = "&System"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "&Start Sip2 Server"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "S&top Sip2 Server"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 3
        Me.MenuItem5.Text = "&About"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 4
        Me.MenuItem6.Text = "&Exit"
        '
        'txtResult
        '
        Me.txtResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtResult.BackColor = System.Drawing.Color.White
        Me.txtResult.ForeColor = System.Drawing.SystemColors.Desktop
        Me.txtResult.Location = New System.Drawing.Point(0, 0)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ReadOnly = True
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtResult.Size = New System.Drawing.Size(704, 336)
        Me.txtResult.TabIndex = 1
        Me.txtResult.TabStop = False
        Me.txtResult.Text = ""
        '
        'tmSip2
        '
        Me.tmSip2.Interval = 1000
        '
        'sbSip2
        '
        Me.sbSip2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.sbSip2.Dock = System.Windows.Forms.DockStyle.None
        Me.sbSip2.Location = New System.Drawing.Point(0, 339)
        Me.sbSip2.Name = "sbSip2"
        Me.sbSip2.Size = New System.Drawing.Size(704, 16)
        Me.sbSip2.TabIndex = 2
        Me.sbSip2.Text = "Sip2Server is running..."
        '
        'frmManagement
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(704, 357)
        Me.Controls.Add(Me.sbSip2)
        Me.Controls.Add(Me.txtResult)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.mmSip2
        Me.Name = "frmManagement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sip2Server service Manager"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private blnStartSip As Boolean
    Private svcSip2Server As ServiceController
    Private isServiceInstalled As Boolean

    Private Sub ShowAbout()
        Dim objAbout As New frmAbout
        objAbout.ShowDialog()
        objAbout.Dispose()
        objAbout = Nothing
    End Sub

    Private Sub CheckServiceInstallation(Optional ByVal blnChangeControl As Boolean = True)
        ' Verify to see if the service is installed.
        Dim installedServices() As ServiceController
        Dim tmpService As ServiceController
        Dim i As Integer = 0

        ' Shut off the timer so it doesn't raise events while were
        '   in this code
        tmSip2.Enabled = False
        installedServices = ServiceController.GetServices()
        isServiceInstalled = False
        For Each tmpService In installedServices
            If tmpService.ServiceName = "Sip2Service" Then
                isServiceInstalled = True
                svcSip2Server = tmpService
                If svcSip2Server.Status = ServiceControllerStatus.Running Then
                    If blnChangeControl Then
                        blnStartSip = True
                        mmSip2.MenuItems(0).MenuItems(0).Enabled = False
                        mmSip2.MenuItems(0).MenuItems(1).Enabled = True
                        cmSip2.MenuItems(1).Enabled = False
                        cmSip2.MenuItems(2).Enabled = True
                    End If
                Else
                    If blnChangeControl Then
                        mmSip2.MenuItems(0).MenuItems(0).Enabled = True
                        mmSip2.MenuItems(0).MenuItems(1).Enabled = False
                        cmSip2.MenuItems(1).Enabled = True
                        cmSip2.MenuItems(2).Enabled = False
                        ntiSip2.Text = "Sip2Server service is stopped..."
                        sbSip2.Text = "Sip2Server service is stopped"
                        blnStartSip = False
                    End If
                End If
                Exit For
            End If
        Next tmpService
        tmSip2.Enabled = True
    End Sub

    Private Sub ChangeStatusSip2Server()
        Me.Cursor = Cursors.WaitCursor
        txtResult.Cursor = Cursors.WaitCursor
        sbSip2.Cursor = Cursors.WaitCursor
        Dim i As Integer
        If blnStartSip Then
            Try
                svcSip2Server.Refresh()
                svcSip2Server.Stop()
                sbSip2.Text = "Sip2Server service is stopping..."
                Application.DoEvents()
                svcSip2Server.Refresh()
                i = 0
                While (Not svcSip2Server.Status = ServiceControllerStatus.Stopped) And (i < 1000)
                    i = i + 1
                    svcSip2Server.Refresh()
                End While
            Catch ex As Exception

            End Try
            If svcSip2Server.Status = ServiceControllerStatus.Stopped Then
                mmSip2.MenuItems(0).MenuItems(0).Enabled = True
                mmSip2.MenuItems(0).MenuItems(1).Enabled = False
                cmSip2.MenuItems(1).Enabled = True
                cmSip2.MenuItems(2).Enabled = False
                ntiSip2.Text = "Sip2Server service is stopped"
                blnStartSip = False
                sbSip2.Text = "Sip2Server service is stopped"
                Application.DoEvents()
            Else
                sbSip2.Text = "Sip2Server service is running..."
                Application.DoEvents()
                MessageBox.Show("Has an error while to try stop service !", "Libol - Sip2Server", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            End If
        Else
            Try
                Dim k As Integer
                For k = 0 To 100
                    svcSip2Server.Refresh()
                    svcSip2Server.Start()
                    sbSip2.Text = "Sip2Server service is starting..."
                    Application.DoEvents()
                    svcSip2Server.Refresh()
                    i = 0
                    While (Not svcSip2Server.Status = ServiceControllerStatus.Running) And (i < 10000)
                        i = i + 1
                        svcSip2Server.Refresh()
                    End While
                    If svcSip2Server.Status = ServiceControllerStatus.Running Then
                        Exit For
                    End If
                Next

            Catch ex As Exception

            End Try
            If svcSip2Server.Status = ServiceControllerStatus.Running Then
                mmSip2.MenuItems(0).MenuItems(0).Enabled = False
                mmSip2.MenuItems(0).MenuItems(1).Enabled = True
                cmSip2.MenuItems(1).Enabled = False
                cmSip2.MenuItems(2).Enabled = True
                ntiSip2.Text = "Sip2Server service is running..."
                blnStartSip = True
                sbSip2.Text = "Sip2Server service is running..."
                Application.DoEvents()
            Else
                sbSip2.Text = "Sip2Server service is stopped"
                Application.DoEvents()
                MessageBox.Show("Has an error while to try start service !", "Libol - Sip2Server", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            End If
        End If
        Me.Cursor = Cursors.Default
        txtResult.Cursor = Cursors.Default
        sbSip2.Cursor = Cursors.Default
    End Sub
    Private Sub GetDataFromToServer()
        Try
            Dim fsLog As StreamReader
            Dim strMessage As String
            fsLog = File.OpenText(Application.StartupPath & "\Sip2Server.log")
            strMessage = fsLog.ReadToEnd()
            fsLog.Close()
            fsLog = Nothing
            If strMessage <> txtResult.Text Then
                txtResult.Text = strMessage
                txtResult.SelectionStart = strMessage.Length
                Application.DoEvents()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Me.Close()
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Call ShowAbout()
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Call ChangeStatusSip2Server()
    End Sub

    Private Sub frmManagement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call CheckServiceInstallation()
        If Not isServiceInstalled Then
            MsgBox("Service Sip2Server is not installed", , "Service is not installed")
            Me.Close()
            Exit Sub
        End If
        ntiSip2.Visible = True
        Me.Visible = False
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Call ChangeStatusSip2Server()
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Call ChangeStatusSip2Server()
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        Call ChangeStatusSip2Server()
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Call ShowAbout()
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Me.Close()
    End Sub

    Private Sub frmManagement_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Visible = False
            ntiSip2.Visible = True
        End If
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        Me.Visible = True
        ntiSip2.Visible = False
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub tmSip2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmSip2.Tick
        Call GetDataFromToServer()
    End Sub

    Private Sub ntiSip2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntiSip2.DoubleClick
        Me.Visible = True
        ntiSip2.Visible = False
        Me.WindowState = FormWindowState.Normal
    End Sub
End Class
