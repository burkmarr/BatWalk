<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectTrack
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectTrack))
        Me.lvTracks = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.Label1 = New System.Windows.Forms.Label
        Me.butCancel = New System.Windows.Forms.Button
        Me.butOkay = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lvTracks
        '
        Me.lvTracks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvTracks.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.lvTracks.FullRowSelect = True
        Me.lvTracks.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvTracks.HideSelection = False
        Me.lvTracks.Location = New System.Drawing.Point(3, 27)
        Me.lvTracks.MultiSelect = False
        Me.lvTracks.Name = "lvTracks"
        Me.lvTracks.Size = New System.Drawing.Size(638, 160)
        Me.lvTracks.TabIndex = 0
        Me.lvTracks.UseCompatibleStateImageBehavior = False
        Me.lvTracks.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Valid track"
        Me.ColumnHeader1.Width = 73
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Start time"
        Me.ColumnHeader2.Width = 66
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "End time"
        Me.ColumnHeader3.Width = 163
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Points"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Track index"
        Me.ColumnHeader5.Width = 73
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(264, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select the track from the GPX file that you wish to use:"
        '
        'butCancel
        '
        Me.butCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butCancel.Location = New System.Drawing.Point(501, 191)
        Me.butCancel.Name = "butCancel"
        Me.butCancel.Size = New System.Drawing.Size(67, 25)
        Me.butCancel.TabIndex = 9
        Me.butCancel.Text = "Cancel"
        Me.butCancel.UseVisualStyleBackColor = True
        '
        'butOkay
        '
        Me.butOkay.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butOkay.Location = New System.Drawing.Point(574, 191)
        Me.butOkay.Name = "butOkay"
        Me.butOkay.Size = New System.Drawing.Size(67, 25)
        Me.butOkay.TabIndex = 8
        Me.butOkay.Text = "Okay"
        Me.butOkay.UseVisualStyleBackColor = True
        '
        'frmSelectTrack
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(646, 218)
        Me.Controls.Add(Me.butCancel)
        Me.Controls.Add(Me.butOkay)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lvTracks)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectTrack"
        Me.Text = "Select track"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvTracks As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents butCancel As System.Windows.Forms.Button
    Friend WithEvents butOkay As System.Windows.Forms.Button
End Class
