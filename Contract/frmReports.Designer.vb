<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReports
    Inherits System.Windows.Forms.Form

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.crvBasic = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'crvBasic
        '
        Me.crvBasic.ActiveViewIndex = -1
        Me.crvBasic.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvBasic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvBasic.Cursor = System.Windows.Forms.Cursors.Default
        Me.crvBasic.Location = New System.Drawing.Point(12, 35)
        Me.crvBasic.Name = "crvBasic"
        Me.crvBasic.Size = New System.Drawing.Size(752, 486)
        Me.crvBasic.TabIndex = 0
        Me.crvBasic.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'frmReports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(768, 533)
        Me.Controls.Add(Me.crvBasic)
        Me.Name = "frmReports"
        Me.Text = "frmReports"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crvBasic As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
