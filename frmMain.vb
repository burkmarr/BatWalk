Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Serialization
Imports System.IO
Imports System.Math
Imports Microsoft.Win32

Public Class frmMain

    Dim bLastTabWasOptions As Boolean
    Dim strOpenedFile As String = ""
    Dim ieCurrent As IntervalExpression

    Public Enum RecDateTimeType
        RecDate
        RecTime
    End Enum

    Public Enum ExportFormat
        R6
        G21
        KML
        Native
    End Enum

    Public Enum IntervalExpression
        HMS
        MS
        S
    End Enum

    Dim dtGPX As DataTable = Nothing

    Private Sub butBrowseGPX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butBrowseGPX.Click

        OpenFileDialog.Filter = "GPX file (*.gpx)|*.gpx|All files (*.*)|*.*"
        OpenFileDialog.FilterIndex = 1

        If Not OpenFileDialog.ShowDialog() = DialogResult.Cancel Then

            txtGPXFile.Text = OpenFileDialog.FileName
            LoadGPX(txtGPXFile.Text)
        End If
    End Sub

    Public Function GetEmptyGPXDatatable() As DataTable

        dtGPX = New DataTable
        dtGPX.Columns.Add("lat")
        dtGPX.Columns.Add("lon")
        dtGPX.Columns.Add("date")
        dtGPX.Columns.Add("time")
        Dim col As DataColumn = New DataColumn("interval", System.Type.GetType("System.Int32"))
        dtGPX.Columns.Add(col)
        Return dtGPX
    End Function
    Private Sub LoadGPX(ByVal sFile As String)

        Dim doc As New XmlDocument()
        Dim ndes As XmlNodeList
        Dim nde As XmlNode
        Dim attr As XmlAttribute
        Dim nde2 As XmlNode
        Dim strLat As String = ""
        Dim strLon As String = ""
        Dim strTime As String = ""
        Dim bStartFound As Boolean = False
        Dim dtmStart As DateTime = Nothing
        Dim dtmThis As DateTime
        Dim tsDiff As TimeSpan
        Dim row As DataRow

        dgvRaw.DataSource = Nothing

        dtGPX = GetEmptyGPXDatatable()
        doc.Load(txtGPXFile.Text)
        frmSelectTrack.gpx = doc
        frmSelectTrack.ShowDialog()
        If frmSelectTrack.TrackSegment Is Nothing Then
            MessageBox.Show("No valid track segment identified.")
            Exit Sub
        End If

        ndes = frmSelectTrack.TrackSegment.SelectNodes("*[contains(name(), 'trkpt')]")

        For Each nde In ndes

            For Each attr In nde.Attributes
                Select Case attr.Name
                    Case "lat"
                        strLat = attr.Value
                    Case "lon"
                        strLon = attr.Value
                    Case Else
                End Select
            Next

            For Each nde2 In nde.ChildNodes
                If nde2.Name = "time" Then
                    strTime = nde2.InnerText
                    Exit For
                End If
            Next

            If Not bStartFound Then
                dtmStart = DateTime.Parse(strTime)
                bStartFound = True
            Else
                dtmThis = DateTime.Parse(strTime)
                tsDiff = dtmThis - dtmStart
            End If

            row = dtGPX.NewRow()
            row("lon") = Convert.ToDouble(strLon)
            row("lat") = Convert.ToDouble(strLat)
            row("date") = UTC2LocalTime(dtmThis, RecDateTimeType.RecDate)
            row("time") = UTC2LocalTime(dtmThis, RecDateTimeType.RecTime)
            row("interval") = tsDiff.TotalSeconds
            dtGPX.Rows.Add(row)

            dgvRaw.DataSource = dtGPX
        Next
    End Sub

    Private Function UTC2LocalTime(ByVal utcDateTime As DateTime, ByVal dtType As RecDateTimeType) As String

        Dim localDateTime As DateTime = utcDateTime.ToLocalTime()

        If dtType = RecDateTimeType.RecDate Then

            Return (Format(localDateTime, "dd/MM/yyyy"))
        Else
            Return (Format(localDateTime, "HH:mm:ss"))
        End If
    End Function

    Private Function GetGridRef(ByVal dblLon As Double, ByVal dblLat As Double, ByVal intRes As Integer)

        Dim dblNorthing As Double
        Dim dblEasting As Double
        Dim strGR As String = ""

        Dim objGridRef As clsGridRef = New clsGridRef
        objGridRef.MakePrefixArrays()
        dblNorthing = objGridRef.LatWGS842Northing(dblLat, dblLon, 100)
        dblEasting = objGridRef.LongWGS842Easting(dblLat, dblLon, 100)

        Select Case intRes
            Case 10000
                strGR = objGridRef.EN2Hectad(dblEasting, dblNorthing)
            Case 2000
                strGR = objGridRef.EN2Tetrad(dblEasting, dblNorthing)
            Case 1000
                strGR = objGridRef.EN2Monad(dblEasting, dblNorthing)
            Case 100
                strGR = objGridRef.EN26fig(dblEasting, dblNorthing)
            Case 10
                strGR = objGridRef.EN28fig(dblEasting, dblNorthing)
            Case 1
                strGR = objGridRef.EN210fig(dblEasting, dblNorthing)
        End Select
        Return strGR
    End Function

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        dgvRecs.Columns.Add("Bat", "Bat")
        dgvRecs.Columns.Add("Abundance", "Abundance")
        dgvRecs.Columns.Add("TreatAs", "TreatAs")
        dgvRecs.Columns.Add("Lat", "Lat")
        dgvRecs.Columns.Add("Lon", "Lon")
        dgvRecs.Columns.Add("GridRef", "GridRef")
        dgvRecs.Columns.Add("DistFromPrev", "DistFromPrev")
        dgvRecs.Columns.Add("Date", "Date")
        dgvRecs.Columns.Add("Time", "Time")
        dgvRecs.Columns.Add("Interval", "Interval")
        dgvRecs.Columns.Add("IntervalSecs", "IntervalSecs")
        dgvRecs.Columns.Add("Comment", "Comment")

        'Set the type and visibility of the IntervalSecs column
        dgvRecs.Columns("IntervalSecs").ValueType = GetType(Integer)
        dgvRecs.Columns("IntervalSecs").Visible = False

        'Select the GPS file tab
        tabMain.SelectedIndex = 1

        ReadRegistrySettings()
    End Sub

    Private Sub butAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butAdd.Click

        If cbPeakFreq.Checked And nudPeakFreq.Value = 0 Then
            MessageBox.Show("You have not set a value for peak frequency. Either set a value or uncheck the box on the options tab for including peak frequency in comment.", "Peak Frequency not set", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If dtGPX Is Nothing Then
            MessageBox.Show("First you must select a GPS track file (GPX file) on the GPS track tab.", "No GPS file selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Dim i As Integer = dgvRecs.Rows.Count
        Dim iTimeInterval As Integer
        Dim foundRows() As DataRow
        Dim strSelectExpr As String
        Dim row As DataRow
        Dim rowClosest As DataRow = Nothing
        Dim iClosestInterval As Integer = 999999
        Dim strComment As String = ""

        'Find point in track that corresponds to the time interval entered
        'Start by selecting all those where time interval is 20 seconds or less 
        'from entered time interval.
        Select Case ieCurrent
            Case IntervalExpression.HMS
                iTimeInterval = nudHours.Value * 3600 + nudMinutes.Value * 60 + nudSeconds.Value
            Case IntervalExpression.MS
                iTimeInterval = nudMinutes.Value * 60 + nudSeconds.Value
            Case IntervalExpression.S
                iTimeInterval = nudSeconds.Value
        End Select

        strSelectExpr = "interval - " & iTimeInterval & " < 20"
        foundRows = dtGPX.Select(strSelectExpr)
        'Go through the selected rows to get the one which is closest to entered interval
        For Each row In foundRows
            If Abs(row("interval") - iTimeInterval) < iClosestInterval Then
                iClosestInterval = Abs(row("interval") - iTimeInterval)
                rowClosest = row
            End If
        Next

        dgvRecs.EditMode = DataGridViewEditMode.EditOnEnter
        dgvRecs.Rows.Insert(i, 1)

        'Initialise values
        dgvRecs.Rows(i).Cells("Date").Value = rowClosest("date")
        dgvRecs.Rows(i).Cells("Time").Value = rowClosest("time")
        dgvRecs.Rows(i).Cells("Interval").Value = ChangeIntervalExpression(iTimeInterval.ToString(), IntervalExpression.S, ieCurrent)
        dgvRecs.Rows(i).Cells("IntervalSecs").Value = Convert.ToInt32(ChangeIntervalExpression(iTimeInterval.ToString(), IntervalExpression.S, ieCurrent))
        dgvRecs.Rows(i).Cells("Lat").Value = rowClosest("lat")
        dgvRecs.Rows(i).Cells("Lon").Value = rowClosest("lon")
        dgvRecs.Rows(i).Cells("GridRef").Value = GetGridRef(rowClosest("lon"), rowClosest("lat"), 10)
        'Comment
        If cbPeakFreq.Checked Then
            strComment = "Peak frequency: " & nudPeakFreq.Value.ToString() & ". "
        End If
        strComment = strComment & txtStandardComment.Text.Trim
        dgvRecs.Rows(i).Cells("Comment").Value = strComment

        dgvRecs.EditMode = DataGridViewEditMode.EditProgrammatically

        'Unselect all records
        Dim dgRow As DataGridViewRow
        For Each dgRow In dgvRecs.Rows
            dgRow.Selected = False
        Next
        'Select the new row
        dgvRecs.Rows(i).Selected = True

        'Put the new row in the right location depending on the interval
        dgvRecs.Sort(dgvRecs.Columns("IntervalSecs"), System.ComponentModel.ListSortDirection.Ascending)

        ResetSeparation()

        'Reset peak frequency
        nudPeakFreq.Value = 0
    End Sub

    Private Sub butDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDelete.Click

        Dim row As DataGridViewRow

        If dgvRecs.SelectedRows.Count = 0 Then

            MessageBox.Show("First you must select the row you want to delete.", "No row selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            If MessageBox.Show("Are you sure that you want to remove this row?", "Confirm deletion", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                For Each row In dgvRecs.SelectedRows
                    dgvRecs.Rows.Remove(row)
                Next

                ResetSeparation()
            End If
        End If
    End Sub

    Private Sub dgvRecs_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvRecs.CellClick

        If e.RowIndex = -1 Then Exit Sub 'Clicked on header
        If e.ColumnIndex = -1 Then Exit Sub 'Selected entire row

        Dim row As DataGridViewRow = dgvRecs.Rows(e.RowIndex)
        Dim strName As String

        If dgvRecs.Columns(e.ColumnIndex).Name = "Bat" Or _
            dgvRecs.Columns(e.ColumnIndex).Name = "Abundance" Or _
            dgvRecs.Columns(e.ColumnIndex).Name = "TreatAs" Or _
            dgvRecs.Columns(e.ColumnIndex).Name = "Comment" Then

            frmRecordProperties.cbBat.Text = row.Cells("Bat").Value
            frmRecordProperties.txtAbundance.Text = row.Cells("Abundance").Value
            If row.Cells("TreatAs").Value = "" Or _
                row.Cells("TreatAs").Value = "Record" Then
                frmRecordProperties.cbTreatAs.SelectedIndex = 0
            Else
                frmRecordProperties.cbTreatAs.SelectedIndex = 1
            End If
            frmRecordProperties.txtComment.Text = row.Cells("Comment").Value

            frmRecordProperties.cbBat.Items.Clear()
            For Each strName In lbBatNames.Items
                frmRecordProperties.cbBat.Items.Add(strName)
            Next

            frmRecordProperties.ShowDialog()

            If frmRecordProperties.Okayed Then
                row.Cells("Bat").Value = frmRecordProperties.cbBat.Text
                row.Cells("Abundance").Value = frmRecordProperties.txtAbundance.Text
                row.Cells("TreatAs").Value = frmRecordProperties.cbTreatAs.Text
                row.Cells("Comment").Value = frmRecordProperties.txtComment.Text
            End If
        End If
    End Sub

    Public Sub WriteRegistrySettings()
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software", True)
        Dim newkey As RegistryKey = key.CreateSubKey("BatWalk")

        Dim strVal As String = ""
        If rbIntHMS.Checked Then strVal = "HMS"
        If rbIntMS.Checked Then strVal = "MS"
        If rbIntS.Checked Then strVal = "S"

        newkey.SetValue("IntervalExpression", strVal)
        newkey.SetValue("EnterPeakFreq", cbPeakFreq.Checked.ToString.ToLower)
        newkey.SetValue("StandardComment", txtStandardComment.Text)
        newkey.SetValue("ExportComment", txtExportComment.Text)

        newkey.DeleteSubKeyTree("BatNames")
        Dim batkey As RegistryKey = newkey.CreateSubKey("BatNames")

        Dim i As Integer
        If lbBatNames.Items.Count > 0 Then
            For i = 1 To lbBatNames.Items.Count
                batkey.SetValue(i.ToString, lbBatNames.Items(i - 1))
            Next
        End If
    End Sub

    Public Sub ReadRegistrySettings()
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software", True)
        Dim newkey As RegistryKey = key.CreateSubKey("BatWalk")
        Dim batkey As RegistryKey = newkey.CreateSubKey("BatNames")

        Dim strVal As String = newkey.GetValue("IntervalExpression")
        Select Case strVal
            Case "MS"
                rbIntMS.Checked = True
            Case "S"
                rbIntS.Checked = True
            Case Else
                rbIntHMS.Checked = True
        End Select

        cbPeakFreq.Checked = (newkey.GetValue("EnterPeakFreq") = "true")
        txtStandardComment.Text = newkey.GetValue("StandardComment")
        txtExportComment.Text = newkey.GetValue("ExportComment")

        Dim i As Integer
        For i = 1 To batkey.ValueCount
            lbBatNames.Items.Add(batkey.GetValue(i.ToString))
        Next
    End Sub

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        WriteRegistrySettings()
    End Sub

    Private Sub tabMain_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabMain.SelectedIndexChanged

        If bLastTabWasOptions Then
            WriteRegistrySettings()
        End If

        If tabMain.SelectedTab.Text = "Options" Then
            bLastTabWasOptions = True
        Else
            bLastTabWasOptions = False
        End If
    End Sub

    Private Sub butNameDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butNameDefaults.Click

        AddBatName("Common Pipistrelle")
        AddBatName("Soprano Pipistrelle")
        AddBatName("Unidentified Pipistrelle")
        AddBatName("Unidentified Myotis")
        AddBatName("Myotis sp.")
        AddBatName("Daubenton's Bat")
        AddBatName("Natterer's Bat")
        AddBatName("Noctule")
        AddBatName("Brown Long-eared Bat")
        AddBatName("Brandt's/Whiskered Bat")
    End Sub

    Private Sub AddBatName(ByVal strName As String)
        If Not lbBatNames.Items.Contains(strName) Then
            lbBatNames.Items.Add(strName)
        End If
    End Sub

    Private Sub butNameAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butNameAdd.Click

        frmNewBatName.ShowDialog()
        If frmNewBatName.Okayed Then
            AddBatName(frmNewBatName.txtBatName.Text)
        End If
    End Sub

    Private Sub butNameRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butNameRemove.Click

        If lbBatNames.SelectedIndex = -1 Then

            MessageBox.Show("First select a bat name to remove from the list.", "No selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            lbBatNames.Items.Remove(lbBatNames.SelectedItem)
        End If
    End Sub

    Private Sub butNameUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butNameUp.Click

        If lbBatNames.SelectedIndex = -1 Then

            MessageBox.Show("First select the bat name you want to move.", "No selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            Dim iSelIndex As Integer = lbBatNames.SelectedIndex

            If iSelIndex > 0 Then
                Dim strName As String = lbBatNames.SelectedItem

                lbBatNames.Items.Insert(iSelIndex - 1, strName)
                lbBatNames.Items.RemoveAt(iSelIndex + 1)
                lbBatNames.SelectedIndex = iSelIndex - 1
            End If
        End If
    End Sub

    Private Sub butNameDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butNameDown.Click

        If lbBatNames.SelectedIndex = -1 Then

            MessageBox.Show("First select the bat name you want to move.", "No selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            Dim iSelIndex As Integer = lbBatNames.SelectedIndex

            If iSelIndex < lbBatNames.Items.Count - 1 Then
                Dim strName As String = lbBatNames.SelectedItem

                If iSelIndex = lbBatNames.Items.Count - 1 Then
                    lbBatNames.Items.Add(strName)
                Else
                    lbBatNames.Items.Insert(iSelIndex + 2, strName)
                End If

                lbBatNames.Items.RemoveAt(iSelIndex)
                lbBatNames.SelectedIndex = iSelIndex + 1
            End If
        End If
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click


    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click

        If strOpenedFile.Length > 0 Then
            SaveFile(strOpenedFile)
        Else
            SaveAsToolStripMenuItem_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click

        SaveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        SaveFileDialog.FilterIndex = 1

        If dgvRecs.RowCount > 0 Then
            If SaveFileDialog.ShowDialog() = DialogResult.OK Then
                SaveFile(SaveFileDialog.FileName)
            End If
        Else
            SaveFile("") 'To invoke the 'no records' messagebox
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click

        If dgvRecs.Rows.Count > 0 Then
            If MessageBox.Show("Are you sure that you want to clear existing rows?", "Confirmation required", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                dgvRecs.Rows.Clear()
                strOpenedFile = ""
            End If
        Else
            strOpenedFile = ""
        End If
    End Sub

    Private Function HasNoValue(ByVal obj As Object) As Boolean

        If IsDBNull(obj) Then

            Return True

        ElseIf obj Is Nothing Then

            Return True

        ElseIf obj.ToString().ToLower = "null" Then

            Return True

        ElseIf obj.ToString.Trim = "" Then

            Return True
        Else
            Return False
        End If

    End Function

    Private Function NullToEmpty(ByVal objValue As Object) As Object

        If HasNoValue(objValue) Then
            Return ""
        Else
            Return objValue
        End If
    End Function

    Private Sub LoadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadToolStripMenuItem.Click

        If dgvRecs.Rows.Count > 0 Then
            Dim response As DialogResult
            response = MessageBox.Show("You already have rows open. Do you want to remove these first(select yes) or append to them (select no)?", "Confirmation required", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If response = DialogResult.Yes Then
                dgvRecs.Rows.Clear()
            ElseIf response = DialogResult.Cancel Then
                Exit Sub
            Else
                'Response was no - so append rows
            End If
        End If

        OpenFileDialog.Filter = "CSV file (*.csv)|*.csv|All files (*.*)|*.*"
        OpenFileDialog.FilterIndex = 1

        If OpenFileDialog.ShowDialog() = DialogResult.Cancel Then
            Exit Sub
        End If

        CreateSchema(OpenFileDialog.FileName)

        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim sFileName As String
        Dim sPathName As String
        Dim sRows As String = ""
        Dim dt As DataTable = New DataTable

        sPathName = System.IO.Path.GetDirectoryName(OpenFileDialog.FileName)
        'Connect to the CSV folder
        MyConnection = New System.Data.OleDb.OleDbConnection( _
             "provider=Microsoft.Jet.OLEDB.4.0; " & _
             "data source=" & sPathName & "; " & _
             "Extended Properties=""text;HDR=Yes;IMEX=1;FMT=Delimited"";")
        MyConnection.Open()

        'Dim conn As New OdbcConnection( _
        '"Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & _
        'Path & ";Extensions=asc,csv,tab,txt")

        'Select the data from relevant CSV file
        sFileName = System.IO.Path.GetFileName(OpenFileDialog.FileName)
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select " & sRows & "* from [" & sFileName & "]", MyConnection)
        MyCommand.Fill(dt)
        MyConnection.Close()

        MyCommand.Dispose()
        MyConnection.Dispose()

        Dim rowDt As DataRow
        Dim rowDgv As DataGridViewRow
        Dim col As DataGridViewColumn
        Dim iRow As Integer


        dgvRecs.EditMode = DataGridViewEditMode.EditOnEnter
        For Each rowDt In dt.Rows
            iRow = dgvRecs.Rows.Add()
            rowDgv = dgvRecs.Rows(iRow)
            For Each col In dgvRecs.Columns
                If dt.Columns.Contains(col.Name) Then

                    If col.Name = "Interval" Then
                        rowDgv.Cells(col.Name).Value = ChangeIntervalExpression(NullToEmpty(rowDt(col.Name)), IntervalExpression.HMS, ieCurrent)
                        rowDgv.Cells("IntervalSecs").Value = Convert.ToInt32(ChangeIntervalExpression(NullToEmpty(rowDt(col.Name)), IntervalExpression.HMS, IntervalExpression.S))
                    Else
                        rowDgv.Cells(col.Name).Value = NullToEmpty(rowDt(col.Name))
                    End If
                Else
                    ''Workaround to populate interval from GPX where not already saved
                    'Dim strVal As String
                    'If col.Name = "Interval" Then
                    '    strVal = WorkaroundGetIntervalFromTime(rowDgv.Cells("Time").Value)
                    '    rowDgv.Cells(col.Name).Value = ChangeIntervalExpression(strVal, IntervalExpression.S, ieCurrent)
                    'End If

                    'If old file with no 'TreatAs' column, put the value 'Record' in
                    If col.Name = "TreatAs" Then
                        rowDgv.Cells(col.Name).Value = "Record"
                    End If
                End If
            Next
        Next
        dgvRecs.EditMode = DataGridViewEditMode.EditProgrammatically

        dgvRecs.Sort(dgvRecs.Columns("IntervalSecs"), System.ComponentModel.ListSortDirection.Ascending)

        ResetSeparation()

        strOpenedFile = OpenFileDialog.FileName
    End Sub

    Private Sub CreateSchema(ByVal sFile As String)

        Dim sFileName As String
        Dim sPathName As String
        Dim sIniFile As String
        Dim sw As StreamWriter
        Dim iCol As Integer = 0
        Dim col As DataColumn
        Dim dt As System.Data.DataTable
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection

        'Delete ini file if already exists
        sPathName = System.IO.Path.GetDirectoryName(sFile)
        sFileName = System.IO.Path.GetFileName(sFile)
        sIniFile = System.IO.Path.Combine(sPathName, "schema.ini")
        If File.Exists(sIniFile) Then
            File.Delete(sIniFile)
        End If

        'Open the csv to get columns 

        MyConnection = New System.Data.OleDb.OleDbConnection( _
                 "provider=Microsoft.Jet.OLEDB.4.0; " & _
                 "data source=" & sPathName & "; " & _
                 "Extended Properties=""text;HDR=Yes;IMEX=1;FMT=Delimited"";")
        MyConnection.Open()

        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & System.IO.Path.GetFileName(sFile) & "]", MyConnection)
        dt = New System.Data.DataTable

        Try
            MyCommand.Fill(dt)
        Catch ex As Exception
            MessageBox.Show("Couldn't open the csv '" & sFile & "'.")
            MyConnection.Close()
            Me.Close()
        End Try

        'Create the schema ini file from the columns
        sw = New StreamWriter(sIniFile)

        sw.WriteLine("[" & sFileName & "]")
        sw.WriteLine("ColNameHeader = True")
        For Each col In dt.Columns
            iCol = iCol + 1


            If iCol = dt.Columns.Count Then
                sw.WriteLine("Col" & iCol.ToString() & "=""" & col.ColumnName & """ Memo")
            Else
                sw.WriteLine("Col" & iCol.ToString() & "=""" & col.ColumnName & """ Text")
            End If
        Next
        sw.Close()

        MyCommand.Dispose()
        MyConnection.Close()
    End Sub

    Private Sub Gilbert21ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Gilbert21ToolStripMenuItem.Click

        ExportRecords(ExportFormat.G21)
    End Sub

    Private Sub ExportRecords(ByVal exform As ExportFormat)

        If dgvRecs.RowCount = 0 Then
            MessageBox.Show("There are no records to export.", "No records", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Dim strExt As String = ""

        Select Case exform
            Case ExportFormat.G21
                strExt = "txt"
            Case ExportFormat.KML
                strExt = "kml"
            Case Else
                strExt = "csv"
        End Select

        SaveFileDialog.Filter = strExt.ToUpper & "files (*." & strExt & ")|*." & strExt & "|All files (*.*)|*.*"
        SaveFileDialog.FilterIndex = 1

        If SaveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim cur As Cursor = Me.Cursor
            Me.Cursor = Cursors.WaitCursor

            Select Case exform
                Case ExportFormat.G21
                    ExportToG21(SaveFileDialog.FileName)
                Case ExportFormat.KML
                    ExportToKML(SaveFileDialog.FileName)
                Case ExportFormat.R6
                    ExportToR6(SaveFileDialog.FileName)
                Case Else
                    'ExportFormat.Native()
            End Select

            Me.Cursor = cur

            If MessageBox.Show("Export complete. Do you want to view the exported file?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                'Open the exported with the default program
                System.Diagnostics.Process.Start(SaveFileDialog.FileName)
            End If
        End If
    End Sub

    Private Sub ExportToG21(ByVal strFile As String)

        Dim sw As StreamWriter = New StreamWriter(strFile)
        Dim row As DataGridViewRow
        Dim strLine As String = ""
        Dim bExclude As Boolean = False
        'Header
        sw.WriteLine("CommonName,TaxonGroup,Abundance,GridRef,RecDate,RecTime,Comment")

        'Rows
        For Each row In dgvRecs.Rows

            bExclude = False
            If cbExDupContacts.Checked Then
                If Not row.Cells("TreatAs").Value = "Record" Then
                    bExclude = True
                End If
            End If

            If Not bExclude Then
                strLine = """" & NullToEmpty(row.Cells("Bat").Value).ToString().Replace("""", """""") & ""","
                strLine = strLine & """terrestrial mammal"","
                strLine = strLine & """" & NullToEmpty(row.Cells("Abundance").Value).ToString().Replace("""", """""") & ""","
                strLine = strLine & """" & row.Cells("GridRef").Value & ""","
                strLine = strLine & """" & row.Cells("Date").Value & ""","
                strLine = strLine & """" & row.Cells("Time").Value.ToString.Substring(0, 5) & ""","
                'strLine = strLine & """" & NullToEmpty(row.Cells("Comment").Value).ToString().Replace("""", """""") & """"
                strLine = strLine & """" & GetExportComment(row.Cells("Comment").Value) & """"
                sw.WriteLine(strLine)
            End If
        Next

        sw.Close()
    End Sub

    Private Function GetExportComment(ByVal strComment As String) As String

        Dim strRet As String
        strRet = NullToEmpty(strComment.Replace("""", """"""))
        If txtExportComment.Text.Trim.Length > 0 Then
            strRet = strRet & " " & txtExportComment.Text.Trim
        End If

        Return strRet
    End Function

    Private Sub ExportToKML(ByVal strFile As String)

        'Create the KML file with the place markers
        Dim srKML As StreamWriter = New StreamWriter(SaveFileDialog.FileName)

        srKML.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
        srKML.WriteLine("<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"">")

        srKML.WriteLine("<Folder>")
        srKML.WriteLine("<name>BatWalk output</name>")

        'Styles for records
        clsKML.StyleKML(srKML, Color.Magenta, "100km")
        clsKML.StyleKML(srKML, Color.Magenta, "hectad")
        clsKML.StyleKML(srKML, Color.Magenta, "tetrad")
        clsKML.StyleKML(srKML, Color.Magenta, "monad")
        clsKML.StyleKML(srKML, Color.Magenta, "6fig")
        clsKML.StyleKML(srKML, Color.Magenta, "8fig")
        clsKML.StyleKML(srKML, Color.Magenta, "10fig")
        clsKML.StyleKML(srKML, Color.Magenta, "invalid")

        'Track point style
        srKML.WriteLine("<Style id=""trackPoint"">")
        srKML.WriteLine("<IconStyle>")
        srKML.WriteLine("<color>ff7fffaa</color>")
        srKML.WriteLine("<scale>0.6</scale>")
        srKML.WriteLine("<Icon>")
        srKML.WriteLine("<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>")
        srKML.WriteLine("</Icon>")
        srKML.WriteLine("</IconStyle>")
        srKML.WriteLine("</Style>")

        'Track line style
        srKML.WriteLine("<Style id=""trackStyle"">")
        srKML.WriteLine("<LineStyle>")
        srKML.WriteLine("<width>2</width>")
        srKML.WriteLine("</LineStyle>")
        srKML.WriteLine("</Style>")

        'Records
        Dim objGridRef As clsGridRef = New clsGridRef
        objGridRef.MakePrefixArrays()
        ExportGERows(srKML, objGridRef)

        'Tracks
        ExportTrackToGE(srKML)

        'Close file
        srKML.WriteLine("</Folder>")
        srKML.WriteLine("</kml>")
        srKML.Close()
    End Sub

    Private Sub ExportGERows(ByVal sw As StreamWriter, ByVal objGridRef As clsGridRef)

        Dim row As DataGridViewRow
        Dim strVal As String
        Dim bExclude As Boolean

        For Each row In dgvRecs.Rows

            bExclude = False
            If cbExDupContacts.Checked Then
                If Not row.Cells("TreatAs").Value = "Record" Then
                    bExclude = True
                End If
            End If

            If Not bExclude Then
                If row.Cells("GridRef").Value.ToString = "" Then
                    Exit Sub
                End If

                sw.WriteLine("<Folder>")
                sw.WriteLine("<name>" & row.Cells("Bat").Value & "</name>")

                sw.WriteLine("<Placemark>")
                sw.WriteLine("<name>" & row.Cells("Bat").Value & "</name>")
                sw.WriteLine("<description>")
                'Values
                Dim strLine As String = ""
                Dim col As DataGridViewColumn
                For Each col In dgvRecs.Columns

                    If col.Visible Then
                        If row.Cells(col.Name).Value Is Nothing Then
                            strVal = ""
                        Else
                            strVal = row.Cells(col.Name).Value.ToString
                        End If

                        If col.Name = "Comment" Then
                            strVal = GetExportComment(strVal)
                        End If
                        strLine = strLine & "<tr><td><b>" & MakeValidXML(col.Name) & "</b></td><td>" & MakeValidXML(NullToEmpty(strVal)) & "</td></tr>"
                    End If
                Next

                objGridRef.GridRef = row.Cells("GridRef").Value.ToString
                objGridRef.ParseGridRef(True)
                objGridRef.ParseInput(False)

                sw.WriteLine("<table>" & strLine & "</table>")
                sw.WriteLine("</description>")
                sw.WriteLine("<styleUrl>#" & objGridRef.sRefType & "</styleUrl>")
                sw.WriteLine("<Point>")
                sw.WriteLine("<coordinates>")

                sw.WriteLine(objGridRef.Easting2LongWGS84(objGridRef.East, objGridRef.North, 0).ToString() & "," & objGridRef.Northing2LatWGS84(objGridRef.East, objGridRef.North, 0).ToString() & ",0")
                sw.WriteLine("</coordinates>")
                sw.WriteLine("</Point>")
                sw.WriteLine("</Placemark>")

                'Render the grid reference squares
                'clsKML.RenderGR(sw, row.Cells("GridRef").Value.ToString)

                sw.WriteLine("</Folder>")
            End If
        Next
    End Sub

    Private Sub ExportToR6(ByVal strFile As String)

        Dim sw As StreamWriter = New StreamWriter(strFile)
        Dim row As DataGridViewRow
        Dim strLine As String = ""
        Dim bExclude As Boolean

        'Header
        sw.WriteLine("Species Name,Abundance Data,Grid Reference,Date,Taxon Occurrence Comment")

        'Rows
        For Each row In dgvRecs.Rows

            bExclude = False
            If cbExDupContacts.Checked Then
                If Not row.Cells("TreatAs").Value = "Record" Then
                    bExclude = True
                End If
            End If

            If Not bExclude Then
                strLine = """" & NullToEmpty(row.Cells("Bat").Value).ToString().Replace("""", """""") & ""","
                strLine = strLine & """" & NullToEmpty(row.Cells("Abundance").Value).ToString().Replace("""", """""") & ""","
                strLine = strLine & """" & row.Cells("GridRef").Value & ""","
                strLine = strLine & """" & row.Cells("Date").Value & ""","
                'strLine = strLine & """" & NullToEmpty(row.Cells("Comment").Value).ToString().Replace("""", """""") & """"
                strLine = strLine & """" & GetExportComment(row.Cells("Comment").Value) & """"
                sw.WriteLine(strLine)
            End If
        Next

        sw.Close()
    End Sub

    Private Sub SaveFile(ByVal strFile As String)

        If dgvRecs.RowCount = 0 Then
            MessageBox.Show("There are no records to save.", "No records", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        Dim strValue As String

        Dim sw As StreamWriter = New StreamWriter(strFile)

        Dim col As DataGridViewColumn
        Dim row As DataGridViewRow

        Dim strLine As String = ""

        'Make the 'DistFromPrev' col invisible so that it's not output
        dgvRecs.Columns("DistFromPrev").Visible = False

        'Header
        For Each col In dgvRecs.Columns
            If col.Visible Then
                If strLine.Length > 0 Then
                    strLine = strLine & ","
                End If
                strLine = strLine & col.Name
            End If
        Next
        sw.WriteLine(strLine)

        'Rows
        For Each row In dgvRecs.Rows
            strLine = ""
            For Each col In dgvRecs.Columns

                If col.Visible Then
                    If strLine.Length > 0 Then
                        strLine = strLine & ","
                    End If

                    If col.Name = "Interval" Then
                        strValue = ChangeIntervalExpression(row.Cells(col.Name).Value, ieCurrent, IntervalExpression.HMS)
                    Else
                        strValue = NullToEmpty(row.Cells(col.Name).Value).ToString().Replace("""", """""")
                    End If

                    strLine = strLine & """" & strValue & """"
                End If
            Next
            sw.WriteLine(strLine)
        Next

        sw.Close()

        'Reset the 'DistFromPrev' col visibility
        dgvRecs.Columns("DistFromPrev").Visible = True

        strOpenedFile = strFile
    End Sub

    Private Sub Recorder6ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Recorder6ToolStripMenuItem.Click

        ExportRecords(ExportFormat.R6)
    End Sub

    Private Sub GoogleEarthToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoogleEarthToolStripMenuItem.Click

        ExportRecords(ExportFormat.KML)
    End Sub
    Private Function MakeValidXML(ByVal objVal As Object) As String

        Dim i As Integer
        Dim chr As Char
        Dim chrNew As String
        Dim strNew As String = ""
        Dim strVal As String = ""

        If Not HasNoValue(objVal) Then
            strVal = objVal.ToString
        End If

        For i = 0 To strVal.Length - 1

            chr = strVal.Substring(i, 1)
            Select Case chr
                Case "&"
                    chrNew = "&amp;"
                Case "<"
                    chrNew = "&lt;"
                Case ">"
                    chrNew = "&gt;"
                Case """"
                    chrNew = "&quot;"
                Case "'"
                    chrNew = "&#39;"
                Case Else
                    'Test for null
                    If Asc(chr) = 0 Then
                        chrNew = ""
                    Else
                        chrNew = chr
                    End If
            End Select

            strNew = strNew & chrNew
        Next

        Return strNew
    End Function

    Private Sub ExportTrackToGE(ByVal sw As StreamWriter)

        If dtGPX Is Nothing Then
            Exit Sub
        End If
        If dtGPX.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim row As DataRow
        sw.WriteLine("<Placemark>")
        sw.WriteLine("<name>BatWalk route</name>")
        sw.WriteLine("<styleUrl>#trackStyle</styleUrl>")
        sw.WriteLine("<LineString>")
        sw.WriteLine("<tessellate>1</tessellate>")
        sw.WriteLine("<coordinates>")

        For Each row In dtGPX.Rows
            sw.WriteLine(row("lon") & "," & row("lat") & ",0 ")
        Next

        sw.WriteLine("</coordinates>")
        sw.WriteLine("</LineString>")
        sw.WriteLine("</Placemark>")
    End Sub

    Private Sub rbIntHMS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbIntHMS.CheckedChanged
        IntervalExpressionChanged()
    End Sub

    Private Sub rbIntMS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbIntMS.CheckedChanged
        IntervalExpressionChanged()
    End Sub

    Private Sub rbIntS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbIntS.CheckedChanged
        IntervalExpressionChanged()
    End Sub

    Private Sub IntervalExpressionChanged()

        nudHours.Value = 0
        nudMinutes.Value = 0

        nudHours.Enabled = False
        nudMinutes.Enabled = False

        If rbIntHMS.Checked Or rbIntMS.Checked Then
            nudMinutes.Enabled = True
        End If

        If rbIntHMS.Checked Then
            nudHours.Enabled = True
        End If

        Dim ieTo As IntervalExpression
        If rbIntHMS.Checked Then ieTo = IntervalExpression.HMS
        If rbIntMS.Checked Then ieTo = IntervalExpression.MS
        If rbIntS.Checked Then ieTo = IntervalExpression.S

        Dim row As DataGridViewRow
        For Each row In dgvRecs.Rows

            row.Cells("Interval").Value = ChangeIntervalExpression(row.Cells("Interval").Value, ieCurrent, ieTo)
        Next

        ieCurrent = ieTo
    End Sub

    Private Function ChangeIntervalExpression(ByVal strValue As String, ByVal ieFrom As IntervalExpression, ByVal ieTo As IntervalExpression) As String

        Dim intHours As Integer
        Dim intMins As Integer
        Dim intSecs As Integer
        Dim intTotalSecs As Integer
        Dim strReturn As String = ""
        Dim strSplit() As String = strValue.Split(":")

        Select Case ieFrom
            Case IntervalExpression.HMS
                intHours = Convert.ToInt32(strSplit(0))
                intMins = Convert.ToInt32(strSplit(1))
                intSecs = Convert.ToInt32(strSplit(2))
                intTotalSecs = intHours * 3600 + intMins * 60 + intSecs
            Case IntervalExpression.MS
                intMins = Convert.ToInt32(strSplit(0))
                intSecs = Convert.ToInt32(strSplit(1))
                intTotalSecs = intMins * 60 + intSecs
            Case IntervalExpression.S
                intTotalSecs = Convert.ToInt32(strValue)
        End Select

        Select Case ieTo
            Case IntervalExpression.HMS
                intHours = intTotalSecs \ 3600
                intMins = (intTotalSecs - 3600 * intHours) \ 60
                intSecs = intTotalSecs - 3600 * intHours - 60 * intMins
                strReturn = intHours.ToString & ":" & intMins.ToString & ":" & intSecs.ToString
            Case IntervalExpression.MS
                intMins = intTotalSecs \ 60
                intSecs = intTotalSecs - 60 * intMins
                strReturn = intMins.ToString & ":" & intSecs.ToString
            Case IntervalExpression.S
                strReturn = intTotalSecs.ToString
        End Select

        Return strReturn
    End Function

    'Private Function WorkaroundGetIntervalFromTime(ByVal strTime As String) As String

    '    Dim foundRows() As DataRow
    '    Dim strSelectExpr As String
    '    Dim row As DataRow
    '    Dim rowClosest As DataRow = Nothing
    '    Dim iClosestInterval As Integer = 999999

    '    'Find point in track that corresponds to the time interval entered
    '    'Start by selecting all those where time interval is 20 seconds or less 
    '    'from entered time interval.
    '    Dim strSplit() As String

    '    strSplit = strTime.Split(":")
    '    Dim iTimeInterval As Integer = Convert.ToInt16(strSplit(0)) * 3600 + Convert.ToInt16(strSplit(1)) * 60 + Convert.ToInt16(strSplit(2))

    '    strSplit = dtGPX.Rows(1)("Time").Split(":")
    '    Dim iStartTimeInterval As Integer = Convert.ToInt16(strSplit(0)) * 3600 + Convert.ToInt16(strSplit(1)) * 60 + Convert.ToInt16(strSplit(2))

    '    iTimeInterval = iTimeInterval - iStartTimeInterval

    '    strSelectExpr = "interval - " & iTimeInterval & " < 20"
    '    foundRows = dtGPX.Select(strSelectExpr)
    '    'Go through the selected rows to get the one which is closest to entered interval
    '    For Each row In foundRows
    '        strSplit = row("Time").Split(":")
    '        If Abs(row("interval") - iTimeInterval) < iClosestInterval Then
    '            iClosestInterval = Abs(row("interval") - iTimeInterval)
    '            rowClosest = row
    '        End If
    '    Next

    '    Return rowClosest("Interval")

    'End Function

    Private Sub ResetSeparation()

        Dim objGridRef As clsGridRef = New clsGridRef
        objGridRef.MakePrefixArrays()
        Dim dblEasting As Double = 0
        Dim dblNorthing As Double = 0
        Dim dblPrevEasting As Double = 0
        Dim dblPrevNorthing As Double = 0

        Dim row As DataGridViewRow
        For Each row In dgvRecs.Rows

            objGridRef.GridRef = row.Cells("GridRef").Value
            objGridRef.ParseGridRef(True)
            objGridRef.ParseInput(False)

            dblEasting = objGridRef.East
            dblNorthing = objGridRef.North

            If dblPrevEasting = 0 Then
                row.Cells("DistFromPrev").Value = 0
            Else
                row.Cells("DistFromPrev").Value = CInt(Sqrt((dblEasting - dblPrevEasting) ^ 2 + (dblNorthing - dblPrevNorthing) ^ 2))
            End If

            dblPrevEasting = dblEasting
            dblPrevNorthing = dblNorthing
        Next
    End Sub
End Class
