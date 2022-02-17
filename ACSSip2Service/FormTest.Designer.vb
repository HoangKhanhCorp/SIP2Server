<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormTest
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lbl = New System.Windows.Forms.Label()
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.chkRequest = New System.Windows.Forms.CheckBox()
        Me.chkResponse = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'lbl
        '
        Me.lbl.AutoSize = True
        Me.lbl.Location = New System.Drawing.Point(12, 9)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(141, 20)
        Me.lbl.TabIndex = 0
        Me.lbl.Text = "Đang khởi động....."
        '
        'txtLog
        '
        Me.txtLog.Location = New System.Drawing.Point(16, 45)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.Size = New System.Drawing.Size(772, 393)
        Me.txtLog.TabIndex = 1
        '
        'chkRequest
        '
        Me.chkRequest.AutoSize = True
        Me.chkRequest.Location = New System.Drawing.Point(16, 454)
        Me.chkRequest.Name = "chkRequest"
        Me.chkRequest.Size = New System.Drawing.Size(150, 24)
        Me.chkRequest.TabIndex = 2
        Me.chkRequest.Text = "Ghi log Request"
        Me.chkRequest.UseVisualStyleBackColor = True
        '
        'chkResponse
        '
        Me.chkResponse.AutoSize = True
        Me.chkResponse.Location = New System.Drawing.Point(405, 454)
        Me.chkResponse.Name = "chkResponse"
        Me.chkResponse.Size = New System.Drawing.Size(162, 24)
        Me.chkResponse.TabIndex = 3
        Me.chkResponse.Text = "Ghi log Response"
        Me.chkResponse.UseVisualStyleBackColor = True
        '
        'FormTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 490)
        Me.Controls.Add(Me.chkResponse)
        Me.Controls.Add(Me.chkRequest)
        Me.Controls.Add(Me.txtLog)
        Me.Controls.Add(Me.lbl)
        Me.Name = "FormTest"
        Me.Text = "Sip2 by ChinhNH"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl As Windows.Forms.Label
    Friend WithEvents txtLog As Windows.Forms.TextBox
    Friend WithEvents chkRequest As Windows.Forms.CheckBox
    Friend WithEvents chkResponse As Windows.Forms.CheckBox
End Class
