Imports System.ComponentModel
Imports System.Configuration.Install

<RunInstaller(True)> Public Class Sip2InstallService
    Inherits System.Configuration.Install.Installer

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Installer overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents myServiceInstaller As System.ServiceProcess.ServiceInstaller
    Friend WithEvents myServiceProcessInstaller As System.ServiceProcess.ServiceProcessInstaller
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.myServiceInstaller = New System.ServiceProcess.ServiceInstaller
        Me.myServiceProcessInstaller = New System.ServiceProcess.ServiceProcessInstaller
        '
        'myServiceProcessInstaller
        '
        Me.myServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.myServiceProcessInstaller.Password = Nothing
        Me.myServiceProcessInstaller.Username = Nothing
        '
        'Sip2InstallService
        '
        Me.myServiceInstaller.ServiceName = "Sip2Service"
        Me.myServiceInstaller.DisplayName = "Sip2Server - Automated Circulation System"
        Me.myServiceInstaller.StartType = ServiceProcess.ServiceStartMode.Automatic
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.myServiceInstaller, Me.myServiceProcessInstaller})

    End Sub

#End Region

End Class
