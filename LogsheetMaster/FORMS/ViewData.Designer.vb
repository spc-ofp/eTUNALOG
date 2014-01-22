<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ViewData
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.CmdInsert = New System.Windows.Forms.Button()
        Me.txtFirstname = New System.Windows.Forms.TextBox()
        Me.txtLastname = New System.Windows.Forms.TextBox()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CmdImportXml = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(38, 28)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(445, 232)
        Me.DataGridView1.TabIndex = 0
        '
        'CmdInsert
        '
        Me.CmdInsert.Location = New System.Drawing.Point(106, 305)
        Me.CmdInsert.Name = "CmdInsert"
        Me.CmdInsert.Size = New System.Drawing.Size(75, 23)
        Me.CmdInsert.TabIndex = 1
        Me.CmdInsert.Text = "Insert"
        Me.CmdInsert.UseVisualStyleBackColor = True
        '
        'txtFirstname
        '
        Me.txtFirstname.Location = New System.Drawing.Point(106, 279)
        Me.txtFirstname.Name = "txtFirstname"
        Me.txtFirstname.Size = New System.Drawing.Size(100, 20)
        Me.txtFirstname.TabIndex = 2
        '
        'txtLastname
        '
        Me.txtLastname.Location = New System.Drawing.Point(212, 279)
        Me.txtLastname.Name = "txtLastname"
        Me.txtLastname.Size = New System.Drawing.Size(100, 20)
        Me.txtLastname.TabIndex = 3
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(318, 279)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(100, 20)
        Me.txtAddress.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(103, 263)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "First Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(209, 263)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Last Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(315, 263)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Address"
        '
        'CmdImportXml
        '
        Me.CmdImportXml.Location = New System.Drawing.Point(106, 347)
        Me.CmdImportXml.Name = "CmdImportXml"
        Me.CmdImportXml.Size = New System.Drawing.Size(138, 23)
        Me.CmdImportXml.TabIndex = 8
        Me.CmdImportXml.Text = "Import XML in database"
        Me.CmdImportXml.UseVisualStyleBackColor = True
        '
        'ViewData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(984, 468)
        Me.Controls.Add(Me.CmdImportXml)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.txtLastname)
        Me.Controls.Add(Me.txtFirstname)
        Me.Controls.Add(Me.CmdInsert)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "ViewData"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents CmdInsert As System.Windows.Forms.Button
    Friend WithEvents txtFirstname As System.Windows.Forms.TextBox
    Friend WithEvents txtLastname As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CmdImportXml As System.Windows.Forms.Button
End Class
