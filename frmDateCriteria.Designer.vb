<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDateCriteria
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtBranch = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDateThru = New System.Windows.Forms.TextBox()
        Me.txtDateFrom = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.cmdButtn01 = New System.Windows.Forms.Button()
        Me.cmdButtn00 = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtBranch)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.txtDateThru)
        Me.Panel1.Controls.Add(Me.txtDateFrom)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ShapeContainer1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(289, 121)
        Me.Panel1.TabIndex = 0
        '
        'txtBranch
        '
        Me.txtBranch.Location = New System.Drawing.Point(73, 17)
        Me.txtBranch.Name = "txtBranch"
        Me.txtBranch.Size = New System.Drawing.Size(203, 20)
        Me.txtBranch.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "BRANCH:"
        '
        'txtDateThru
        '
        Me.txtDateThru.Location = New System.Drawing.Point(90, 88)
        Me.txtDateThru.Name = "txtDateThru"
        Me.txtDateThru.Size = New System.Drawing.Size(187, 20)
        Me.txtDateThru.TabIndex = 3
        '
        'txtDateFrom
        '
        Me.txtDateFrom.BackColor = System.Drawing.Color.White
        Me.txtDateFrom.Location = New System.Drawing.Point(91, 62)
        Me.txtDateFrom.Name = "txtDateFrom"
        Me.txtDateFrom.Size = New System.Drawing.Size(187, 20)
        Me.txtDateFrom.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 91)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "DATE THRU:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "DATE FROM:"
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(287, 119)
        Me.ShapeContainer1.TabIndex = 5
        Me.ShapeContainer1.TabStop = False
        '
        'LineShape1
        '
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 12
        Me.LineShape1.X2 = 277
        Me.LineShape1.Y1 = 53
        Me.LineShape1.Y2 = 53
        '
        'cmdButtn01
        '
        Me.cmdButtn01.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdButtn01.Location = New System.Drawing.Point(317, 48)
        Me.cmdButtn01.Name = "cmdButtn01"
        Me.cmdButtn01.Size = New System.Drawing.Size(79, 32)
        Me.cmdButtn01.TabIndex = 21
        Me.cmdButtn01.Text = "&Cancel"
        Me.cmdButtn01.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButtn01.UseVisualStyleBackColor = True
        '
        'cmdButtn00
        '
        Me.cmdButtn00.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdButtn00.Location = New System.Drawing.Point(317, 14)
        Me.cmdButtn00.Name = "cmdButtn00"
        Me.cmdButtn00.Size = New System.Drawing.Size(79, 32)
        Me.cmdButtn00.TabIndex = 20
        Me.cmdButtn00.Text = "&Ok"
        Me.cmdButtn00.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButtn00.UseVisualStyleBackColor = True
        '
        'frmDateCriteria
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ClientSize = New System.Drawing.Size(408, 142)
        Me.Controls.Add(Me.cmdButtn01)
        Me.Controls.Add(Me.cmdButtn00)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDateCriteria"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Date Criteria"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtDateFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDateThru As System.Windows.Forms.TextBox
    Friend WithEvents cmdButtn01 As System.Windows.Forms.Button
    Friend WithEvents cmdButtn00 As System.Windows.Forms.Button
    Friend WithEvents txtBranch As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineShape1 As Microsoft.VisualBasic.PowerPacks.LineShape
End Class
