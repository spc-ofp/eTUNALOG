<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FirstSteps
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FirstSteps))
        Me.MyRtb1 = New eTunaLog.MyRtb()
        Me.SuspendLayout()
        '
        'MyRtb1
        '
        Me.MyRtb1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyRtb1.Location = New System.Drawing.Point(12, 12)
        Me.MyRtb1.Margin = New System.Windows.Forms.Padding(10, 3, 10, 3)
        Me.MyRtb1.Name = "MyRtb1"
        Me.MyRtb1.ReadOnly = True
        Me.MyRtb1.RichText = resources.GetString("MyRtb1.RichText")
        Me.MyRtb1.Size = New System.Drawing.Size(490, 379)
        Me.MyRtb1.TabIndex = 0
        Me.MyRtb1.Text = resources.GetString("MyRtb1.Text")
        '
        'FirstSteps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(514, 403)
        Me.Controls.Add(Me.MyRtb1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FirstSteps"
        Me.Text = "First steps in eTUNALOG..."
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MyRtb1 As eTunaLog.MyRtb
End Class
