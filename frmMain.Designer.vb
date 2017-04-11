<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.butBrowseGPX = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtGPXFile = New System.Windows.Forms.TextBox
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.dgvRaw = New System.Windows.Forms.DataGridView
        Me.tabMain = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.nudHours = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.butDelete = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.dgvRecs = New System.Windows.Forms.DataGridView
        Me.nudPeakFreq = New System.Windows.Forms.NumericUpDown
        Me.butAdd = New System.Windows.Forms.Button
        Me.nudMinutes = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.nudSeconds = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.cbExDupContacts = New System.Windows.Forms.CheckBox
        Me.rbIntS = New System.Windows.Forms.RadioButton
        Me.rbIntMS = New System.Windows.Forms.RadioButton
        Me.rbIntHMS = New System.Windows.Forms.RadioButton
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtExportComment = New System.Windows.Forms.TextBox
        Me.butNameDefaults = New System.Windows.Forms.Button
        Me.butNameDown = New System.Windows.Forms.Button
        Me.butNameUp = New System.Windows.Forms.Button
        Me.butNameRemove = New System.Windows.Forms.Button
        Me.butNameAdd = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.lbBatNames = New System.Windows.Forms.ListBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtStandardComment = New System.Windows.Forms.TextBox
        Me.cbPeakFreq = New System.Windows.Forms.CheckBox
        Me.MenuStrip = New System.Windows.Forms.MenuStrip
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.LoadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Recorder6ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Gilbert21ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GoogleEarthToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        CType(Me.dgvRaw, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabMain.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.nudHours, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvRecs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudPeakFreq, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMinutes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudSeconds, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.MenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'butBrowseGPX
        '
        Me.butBrowseGPX.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butBrowseGPX.Location = New System.Drawing.Point(758, 6)
        Me.butBrowseGPX.Name = "butBrowseGPX"
        Me.butBrowseGPX.Size = New System.Drawing.Size(53, 20)
        Me.butBrowseGPX.TabIndex = 40
        Me.butBrowseGPX.Text = "Browse"
        Me.butBrowseGPX.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(7, 9)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(48, 13)
        Me.Label11.TabIndex = 41
        Me.Label11.Text = "GPX file:"
        '
        'txtGPXFile
        '
        Me.txtGPXFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGPXFile.Location = New System.Drawing.Point(55, 6)
        Me.txtGPXFile.Name = "txtGPXFile"
        Me.txtGPXFile.Size = New System.Drawing.Size(697, 20)
        Me.txtGPXFile.TabIndex = 39
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog"
        Me.OpenFileDialog.Multiselect = True
        '
        'dgvRaw
        '
        Me.dgvRaw.AllowUserToAddRows = False
        Me.dgvRaw.AllowUserToDeleteRows = False
        Me.dgvRaw.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRaw.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRaw.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvRaw.Location = New System.Drawing.Point(0, 32)
        Me.dgvRaw.Name = "dgvRaw"
        Me.dgvRaw.Size = New System.Drawing.Size(817, 389)
        Me.dgvRaw.TabIndex = 42
        '
        'tabMain
        '
        Me.tabMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabMain.Controls.Add(Me.TabPage1)
        Me.tabMain.Controls.Add(Me.TabPage2)
        Me.tabMain.Controls.Add(Me.TabPage3)
        Me.tabMain.Location = New System.Drawing.Point(3, 28)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SelectedIndex = 0
        Me.tabMain.Size = New System.Drawing.Size(825, 443)
        Me.tabMain.TabIndex = 43
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.nudHours)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.butDelete)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.dgvRecs)
        Me.TabPage1.Controls.Add(Me.nudPeakFreq)
        Me.TabPage1.Controls.Add(Me.butAdd)
        Me.TabPage1.Controls.Add(Me.nudMinutes)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.nudSeconds)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(817, 417)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Bat contacts"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'nudHours
        '
        Me.nudHours.Location = New System.Drawing.Point(38, 7)
        Me.nudHours.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.nudHours.Name = "nudHours"
        Me.nudHours.Size = New System.Drawing.Size(45, 20)
        Me.nudHours.TabIndex = 52
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(0, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(38, 13)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Hours:"
        '
        'butDelete
        '
        Me.butDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butDelete.Location = New System.Drawing.Point(758, 7)
        Me.butDelete.Name = "butDelete"
        Me.butDelete.Size = New System.Drawing.Size(53, 20)
        Me.butDelete.TabIndex = 51
        Me.butDelete.Text = "Delete"
        Me.butDelete.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(354, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 50
        Me.Label3.Text = "Peak freq:"
        '
        'dgvRecs
        '
        Me.dgvRecs.AllowUserToAddRows = False
        Me.dgvRecs.AllowUserToDeleteRows = False
        Me.dgvRecs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRecs.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvRecs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRecs.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvRecs.Location = New System.Drawing.Point(0, 33)
        Me.dgvRecs.Name = "dgvRecs"
        Me.dgvRecs.Size = New System.Drawing.Size(816, 384)
        Me.dgvRecs.TabIndex = 0
        '
        'nudPeakFreq
        '
        Me.nudPeakFreq.DecimalPlaces = 1
        Me.nudPeakFreq.Location = New System.Drawing.Point(410, 7)
        Me.nudPeakFreq.Maximum = New Decimal(New Integer() {120, 0, 0, 0})
        Me.nudPeakFreq.Name = "nudPeakFreq"
        Me.nudPeakFreq.Size = New System.Drawing.Size(61, 20)
        Me.nudPeakFreq.TabIndex = 49
        '
        'butAdd
        '
        Me.butAdd.Location = New System.Drawing.Point(496, 7)
        Me.butAdd.Name = "butAdd"
        Me.butAdd.Size = New System.Drawing.Size(53, 20)
        Me.butAdd.TabIndex = 48
        Me.butAdd.Text = "Add"
        Me.butAdd.UseVisualStyleBackColor = True
        '
        'nudMinutes
        '
        Me.nudMinutes.Location = New System.Drawing.Point(127, 7)
        Me.nudMinutes.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.nudMinutes.Name = "nudMinutes"
        Me.nudMinutes.Size = New System.Drawing.Size(57, 20)
        Me.nudMinutes.TabIndex = 44
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(196, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 47
        Me.Label2.Text = "Secs:"
        '
        'nudSeconds
        '
        Me.nudSeconds.Location = New System.Drawing.Point(230, 7)
        Me.nudSeconds.Maximum = New Decimal(New Integer() {20000, 0, 0, 0})
        Me.nudSeconds.Name = "nudSeconds"
        Me.nudSeconds.Size = New System.Drawing.Size(81, 20)
        Me.nudSeconds.TabIndex = 45
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(95, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 46
        Me.Label1.Text = "Mins:"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvRaw)
        Me.TabPage2.Controls.Add(Me.txtGPXFile)
        Me.TabPage2.Controls.Add(Me.Label11)
        Me.TabPage2.Controls.Add(Me.butBrowseGPX)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(817, 417)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "GPS track"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.cbExDupContacts)
        Me.TabPage3.Controls.Add(Me.rbIntS)
        Me.TabPage3.Controls.Add(Me.rbIntMS)
        Me.TabPage3.Controls.Add(Me.rbIntHMS)
        Me.TabPage3.Controls.Add(Me.Label8)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Controls.Add(Me.txtExportComment)
        Me.TabPage3.Controls.Add(Me.butNameDefaults)
        Me.TabPage3.Controls.Add(Me.butNameDown)
        Me.TabPage3.Controls.Add(Me.butNameUp)
        Me.TabPage3.Controls.Add(Me.butNameRemove)
        Me.TabPage3.Controls.Add(Me.butNameAdd)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.lbBatNames)
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.txtStandardComment)
        Me.TabPage3.Controls.Add(Me.cbPeakFreq)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(817, 417)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Options"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'cbExDupContacts
        '
        Me.cbExDupContacts.AutoSize = True
        Me.cbExDupContacts.Checked = True
        Me.cbExDupContacts.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbExDupContacts.Location = New System.Drawing.Point(229, 31)
        Me.cbExDupContacts.Name = "cbExDupContacts"
        Me.cbExDupContacts.Size = New System.Drawing.Size(209, 17)
        Me.cbExDupContacts.TabIndex = 16
        Me.cbExDupContacts.Text = "Exclude duplicate contacts from export"
        Me.cbExDupContacts.UseVisualStyleBackColor = True
        '
        'rbIntS
        '
        Me.rbIntS.AutoSize = True
        Me.rbIntS.Location = New System.Drawing.Point(393, 8)
        Me.rbIntS.Name = "rbIntS"
        Me.rbIntS.Size = New System.Drawing.Size(89, 17)
        Me.rbIntS.TabIndex = 15
        Me.rbIntS.TabStop = True
        Me.rbIntS.Text = "Seconds only"
        Me.rbIntS.UseVisualStyleBackColor = True
        '
        'rbIntMS
        '
        Me.rbIntMS.AutoSize = True
        Me.rbIntMS.Location = New System.Drawing.Point(281, 8)
        Me.rbIntMS.Name = "rbIntMS"
        Me.rbIntMS.Size = New System.Drawing.Size(93, 17)
        Me.rbIntMS.TabIndex = 14
        Me.rbIntMS.TabStop = True
        Me.rbIntMS.Text = "Mins and secs"
        Me.rbIntMS.UseVisualStyleBackColor = True
        '
        'rbIntHMS
        '
        Me.rbIntHMS.AutoSize = True
        Me.rbIntHMS.Location = New System.Drawing.Point(134, 8)
        Me.rbIntHMS.Name = "rbIntHMS"
        Me.rbIntHMS.Size = New System.Drawing.Size(126, 17)
        Me.rbIntHMS.TabIndex = 13
        Me.rbIntHMS.TabStop = True
        Me.rbIntHMS.Text = "Hours, mins and secs"
        Me.rbIntHMS.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 10)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(123, 13)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "Time interval expression:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 137)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(187, 13)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Text to append to comment on export:"
        '
        'txtExportComment
        '
        Me.txtExportComment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExportComment.Location = New System.Drawing.Point(5, 153)
        Me.txtExportComment.Multiline = True
        Me.txtExportComment.Name = "txtExportComment"
        Me.txtExportComment.Size = New System.Drawing.Size(807, 58)
        Me.txtExportComment.TabIndex = 10
        '
        'butNameDefaults
        '
        Me.butNameDefaults.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butNameDefaults.Location = New System.Drawing.Point(727, 374)
        Me.butNameDefaults.Name = "butNameDefaults"
        Me.butNameDefaults.Size = New System.Drawing.Size(83, 36)
        Me.butNameDefaults.TabIndex = 9
        Me.butNameDefaults.Text = "Restore default names"
        Me.butNameDefaults.UseVisualStyleBackColor = True
        '
        'butNameDown
        '
        Me.butNameDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butNameDown.Location = New System.Drawing.Point(727, 333)
        Me.butNameDown.Name = "butNameDown"
        Me.butNameDown.Size = New System.Drawing.Size(83, 22)
        Me.butNameDown.TabIndex = 8
        Me.butNameDown.Text = "Down"
        Me.butNameDown.UseVisualStyleBackColor = True
        '
        'butNameUp
        '
        Me.butNameUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butNameUp.Location = New System.Drawing.Point(727, 305)
        Me.butNameUp.Name = "butNameUp"
        Me.butNameUp.Size = New System.Drawing.Size(83, 22)
        Me.butNameUp.TabIndex = 7
        Me.butNameUp.Text = "Up"
        Me.butNameUp.UseVisualStyleBackColor = True
        '
        'butNameRemove
        '
        Me.butNameRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butNameRemove.Location = New System.Drawing.Point(727, 266)
        Me.butNameRemove.Name = "butNameRemove"
        Me.butNameRemove.Size = New System.Drawing.Size(83, 22)
        Me.butNameRemove.TabIndex = 6
        Me.butNameRemove.Text = "Remove"
        Me.butNameRemove.UseVisualStyleBackColor = True
        '
        'butNameAdd
        '
        Me.butNameAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.butNameAdd.Location = New System.Drawing.Point(727, 238)
        Me.butNameAdd.Name = "butNameAdd"
        Me.butNameAdd.Size = New System.Drawing.Size(83, 22)
        Me.butNameAdd.TabIndex = 5
        Me.butNameAdd.Text = "Add"
        Me.butNameAdd.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 222)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Bat names:"
        '
        'lbBatNames
        '
        Me.lbBatNames.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbBatNames.FormattingEnabled = True
        Me.lbBatNames.Location = New System.Drawing.Point(5, 238)
        Me.lbBatNames.Name = "lbBatNames"
        Me.lbBatNames.Size = New System.Drawing.Size(718, 173)
        Me.lbBatNames.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(119, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Standard comment text:"
        '
        'txtStandardComment
        '
        Me.txtStandardComment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStandardComment.Location = New System.Drawing.Point(5, 67)
        Me.txtStandardComment.Multiline = True
        Me.txtStandardComment.Name = "txtStandardComment"
        Me.txtStandardComment.Size = New System.Drawing.Size(807, 58)
        Me.txtStandardComment.TabIndex = 1
        '
        'cbPeakFreq
        '
        Me.cbPeakFreq.AutoSize = True
        Me.cbPeakFreq.Checked = True
        Me.cbPeakFreq.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbPeakFreq.Location = New System.Drawing.Point(8, 31)
        Me.cbPeakFreq.Name = "cbPeakFreq"
        Me.cbPeakFreq.Size = New System.Drawing.Size(180, 17)
        Me.cbPeakFreq.TabIndex = 0
        Me.cbPeakFreq.Text = "Add peak frequency to comment"
        Me.cbPeakFreq.UseVisualStyleBackColor = True
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(831, 24)
        Me.MenuStrip.TabIndex = 44
        Me.MenuStrip.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.LoadToolStripMenuItem, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.ExportToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(35, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.NewToolStripMenuItem.Text = "New"
        '
        'LoadToolStripMenuItem
        '
        Me.LoadToolStripMenuItem.Name = "LoadToolStripMenuItem"
        Me.LoadToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.LoadToolStripMenuItem.Text = "Open"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.SaveToolStripMenuItem.Text = "Save"
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.SaveAsToolStripMenuItem.Text = "Save as"
        '
        'ExportToolStripMenuItem
        '
        Me.ExportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Recorder6ToolStripMenuItem, Me.Gilbert21ToolStripMenuItem, Me.GoogleEarthToolStripMenuItem})
        Me.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem"
        Me.ExportToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.ExportToolStripMenuItem.Text = "Export"
        '
        'Recorder6ToolStripMenuItem
        '
        Me.Recorder6ToolStripMenuItem.Name = "Recorder6ToolStripMenuItem"
        Me.Recorder6ToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.Recorder6ToolStripMenuItem.Text = "Recorder 6"
        '
        'Gilbert21ToolStripMenuItem
        '
        Me.Gilbert21ToolStripMenuItem.Name = "Gilbert21ToolStripMenuItem"
        Me.Gilbert21ToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.Gilbert21ToolStripMenuItem.Text = "Gilbert 21"
        '
        'GoogleEarthToolStripMenuItem
        '
        Me.GoogleEarthToolStripMenuItem.Name = "GoogleEarthToolStripMenuItem"
        Me.GoogleEarthToolStripMenuItem.Size = New System.Drawing.Size(147, 22)
        Me.GoogleEarthToolStripMenuItem.Text = "Google Earth"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(831, 473)
        Me.Controls.Add(Me.tabMain)
        Me.Controls.Add(Me.MenuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "frmMain"
        Me.Text = "BatWalk"
        CType(Me.dgvRaw, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabMain.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.nudHours, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvRecs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudPeakFreq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMinutes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudSeconds, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents butBrowseGPX As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtGPXFile As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents dgvRaw As System.Windows.Forms.DataGridView
    Friend WithEvents tabMain As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvRecs As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents nudMinutes As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudSeconds As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents butAdd As System.Windows.Forms.Button
    Friend WithEvents nudPeakFreq As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtStandardComment As System.Windows.Forms.TextBox
    Friend WithEvents cbPeakFreq As System.Windows.Forms.CheckBox
    Friend WithEvents butDelete As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lbBatNames As System.Windows.Forms.ListBox
    Friend WithEvents butNameDown As System.Windows.Forms.Button
    Friend WithEvents butNameUp As System.Windows.Forms.Button
    Friend WithEvents butNameRemove As System.Windows.Forms.Button
    Friend WithEvents butNameAdd As System.Windows.Forms.Button
    Friend WithEvents butNameDefaults As System.Windows.Forms.Button
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents LoadToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Recorder6ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Gilbert21ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GoogleEarthToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents nudHours As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtExportComment As System.Windows.Forms.TextBox
    Friend WithEvents rbIntMS As System.Windows.Forms.RadioButton
    Friend WithEvents rbIntHMS As System.Windows.Forms.RadioButton
    Friend WithEvents rbIntS As System.Windows.Forms.RadioButton
    Friend WithEvents cbExDupContacts As System.Windows.Forms.CheckBox

End Class
