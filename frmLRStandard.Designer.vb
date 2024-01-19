<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLRStandard
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLRStandard))
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtField01 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtField00 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rbtDetail = New System.Windows.Forms.RadioButton()
        Me.rbtSummary = New System.Windows.Forms.RadioButton()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.txtField83 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtField82 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtField81 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtField80 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdButtn01 = New System.Windows.Forms.Button()
        Me.cmdButtn00 = New System.Windows.Forms.Button()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Panel2)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Location = New System.Drawing.Point(12, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(388, 79)
        Me.Panel3.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtField01)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.txtField00)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(124, 10)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(253, 57)
        Me.Panel2.TabIndex = 4
        '
        'txtField01
        '
        Me.txtField01.Location = New System.Drawing.Point(101, 28)
        Me.txtField01.Name = "txtField01"
        Me.txtField01.Size = New System.Drawing.Size(140, 20)
        Me.txtField01.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Date Thru"
        '
        'txtField00
        '
        Me.txtField00.Location = New System.Drawing.Point(101, 6)
        Me.txtField00.Name = "txtField00"
        Me.txtField00.Size = New System.Drawing.Size(140, 20)
        Me.txtField00.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(20, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Date From"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.rbtDetail)
        Me.Panel1.Controls.Add(Me.rbtSummary)
        Me.Panel1.Location = New System.Drawing.Point(11, 10)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(107, 57)
        Me.Panel1.TabIndex = 1
        '
        'rbtDetail
        '
        Me.rbtDetail.AutoSize = True
        Me.rbtDetail.Checked = True
        Me.rbtDetail.Location = New System.Drawing.Point(20, 29)
        Me.rbtDetail.Name = "rbtDetail"
        Me.rbtDetail.Size = New System.Drawing.Size(52, 17)
        Me.rbtDetail.TabIndex = 3
        Me.rbtDetail.TabStop = True
        Me.rbtDetail.Text = "Detail"
        Me.rbtDetail.UseVisualStyleBackColor = True
        '
        'rbtSummary
        '
        Me.rbtSummary.AutoSize = True
        Me.rbtSummary.Location = New System.Drawing.Point(20, 6)
        Me.rbtSummary.Name = "rbtSummary"
        Me.rbtSummary.Size = New System.Drawing.Size(68, 17)
        Me.rbtSummary.TabIndex = 2
        Me.rbtSummary.TabStop = True
        Me.rbtSummary.Text = "Summary"
        Me.rbtSummary.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.txtField83)
        Me.Panel4.Controls.Add(Me.Label6)
        Me.Panel4.Controls.Add(Me.txtField82)
        Me.Panel4.Controls.Add(Me.Label5)
        Me.Panel4.Controls.Add(Me.txtField81)
        Me.Panel4.Controls.Add(Me.Label4)
        Me.Panel4.Controls.Add(Me.txtField80)
        Me.Panel4.Controls.Add(Me.Label3)
        Me.Panel4.Location = New System.Drawing.Point(12, 98)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(388, 107)
        Me.Panel4.TabIndex = 9
        '
        'txtField83
        '
        Me.txtField83.Location = New System.Drawing.Point(88, 75)
        Me.txtField83.Name = "txtField83"
        Me.txtField83.Size = New System.Drawing.Size(289, 20)
        Me.txtField83.TabIndex = 17
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 79)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 13)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Collector"
        '
        'txtField82
        '
        Me.txtField82.Location = New System.Drawing.Point(88, 53)
        Me.txtField82.Name = "txtField82"
        Me.txtField82.Size = New System.Drawing.Size(289, 20)
        Me.txtField82.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 57)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(29, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Area"
        '
        'txtField81
        '
        Me.txtField81.Location = New System.Drawing.Point(88, 31)
        Me.txtField81.Name = "txtField81"
        Me.txtField81.Size = New System.Drawing.Size(289, 20)
        Me.txtField81.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(41, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Branch"
        '
        'txtField80
        '
        Me.txtField80.Location = New System.Drawing.Point(88, 9)
        Me.txtField80.Name = "txtField80"
        Me.txtField80.Size = New System.Drawing.Size(289, 20)
        Me.txtField80.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Company"
        '
        'cmdButtn01
        '
        Me.cmdButtn01.Image = CType(resources.GetObject("cmdButtn01.Image"), System.Drawing.Image)
        Me.cmdButtn01.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdButtn01.Location = New System.Drawing.Point(415, 46)
        Me.cmdButtn01.Name = "cmdButtn01"
        Me.cmdButtn01.Size = New System.Drawing.Size(79, 32)
        Me.cmdButtn01.TabIndex = 19
        Me.cmdButtn01.Text = "&Cancel"
        Me.cmdButtn01.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButtn01.UseVisualStyleBackColor = True
        '
        'cmdButtn00
        '
        Me.cmdButtn00.Image = CType(resources.GetObject("cmdButtn00.Image"), System.Drawing.Image)
        Me.cmdButtn00.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdButtn00.Location = New System.Drawing.Point(415, 12)
        Me.cmdButtn00.Name = "cmdButtn00"
        Me.cmdButtn00.Size = New System.Drawing.Size(79, 32)
        Me.cmdButtn00.TabIndex = 18
        Me.cmdButtn00.Text = "&Ok"
        Me.cmdButtn00.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.cmdButtn00.UseVisualStyleBackColor = True
        '
        'frmLRStandard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ClientSize = New System.Drawing.Size(502, 215)
        Me.Controls.Add(Me.cmdButtn01)
        Me.Controls.Add(Me.cmdButtn00)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLRStandard"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Standard Report"
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtField01 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtField00 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rbtDetail As System.Windows.Forms.RadioButton
    Friend WithEvents rbtSummary As System.Windows.Forms.RadioButton
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents txtField82 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtField81 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtField80 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdButtn01 As System.Windows.Forms.Button
    Friend WithEvents cmdButtn00 As System.Windows.Forms.Button
    Friend WithEvents txtField83 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
