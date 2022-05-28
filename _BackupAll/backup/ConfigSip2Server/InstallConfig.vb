Imports System.ComponentModel
Imports System.Configuration.Install

<RunInstaller(True)> Public Class InstallConfig
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Private Sub InstallConfig_BeforeInstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles MyBase.BeforeInstall
        Try
            Dim frmConfigtmp As New frmConfig
            frmConfigtmp.ShowDialog()
            frmConfigtmp.Dispose()
            frmConfigtmp = Nothing
        Catch ex As Exception
        End Try
    End Sub
End Class
