<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReports
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReports))
        Me.dgView = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cmdButtn01 = New System.Windows.Forms.Button()
        Me.cmdButtn00 = New System.Windows.Forms.Button()
        CType(Me.dgView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgView
        '
        Me.dgView.AllowUserToAddRows = False
        Me.dgView.AllowUserToDeleteRows = False
        Me.dgView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1})
        Me.dgView.Location = New System.Drawing.Point(5, 8)
        Me.dgView.Name = "dgView"
        Me.dgView.ReadOnly = True
        Me.dgView.Size = New System.Drawing.Size(414, 418)
        Me.dgView.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.FillWeight = 365.0!
        Me.Column1.HeaderText = "Reports"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Column1.Width = 365
        '
        'cmdButtn01
        '
        Me.cmdButtn01.Image = CType(resources.GetObject("cmdButtn01.Image"), System.Drawing.Image)
        Me.cmdButtn01.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdButtn01.Location = New System.Drawing.Point(436, 100)
        Me.cmdButtn01.Name = "cmdButtn01"
        Me.cmdButtn01.Size = New System.Drawing.Size(79, 32)
        Me.cmdButtn01.TabIndex = 67
        Me.cmdButtn01.Text = "&Cancel"
        Me.cmdButtn01.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButtn01.UseVisualStyleBackColor = True
        '
        'cmdButtn00
        '
        Me.cmdButtn00.Image = CType(resources.GetObject("cmdButtn00.Image"), System.Drawing.Image)
        Me.cmdButtn00.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdButtn00.Location = New System.Drawing.Point(436, 66)
        Me.cmdButtn00.Name = "cmdButtn00"
        Me.cmdButtn00.Size = New System.Drawing.Size(79, 32)
        Me.cmdButtn00.TabIndex = 66
        Me.cmdButtn00.Text = "&Ok"
        Me.cmdButtn00.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButtn00.UseVisualStyleBackColor = True
        '
        'frmReports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(525, 432)
        Me.Controls.Add(Me.cmdButtn01)
        Me.Controls.Add(Me.cmdButtn00)
        Me.Controls.Add(Me.dgView)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReports"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "LR Reports"
        CType(Me.dgView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgView As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cmdButtn01 As System.Windows.Forms.Button
    Friend WithEvents cmdButtn00 As System.Windows.Forms.Button
End Class
