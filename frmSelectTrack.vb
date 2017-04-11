Imports System.Xml
Imports System.Xml.XPath

Public Class frmSelectTrack

    Public Enum FirstLast
        First
        Last
    End Enum

    Public Enum RecDateTimeType
        RecDate
        RecTime
    End Enum

    Private xmlDoc As XmlDocument
    Public Property gpx() As XmlDocument
        Get
            Return xmlDoc
        End Get
        Set(ByVal value As XmlDocument)
            xmlDoc = value
        End Set
    End Property

    Private ndeTrkSeg As XmlNode
    Public ReadOnly Property TrackSegment() As XmlNode
        Get
            Return ndeTrkSeg
        End Get
    End Property

    Dim ndes As XmlNodeList

    Public Sub InitialiseList()

        Dim nde As XmlNode
        Dim ndesTrkPt As XmlNodeList
        Dim strStart As String
        Dim strEnd As String
        Dim iValidTrack As Integer = 0

        ndes = gpx.GetElementsByTagName("trkseg")
        Dim li As ListViewItem
        Dim iTrackSeg As Integer
        Dim str(4) As String

        lvTracks.Items.Clear()
        For iTrackSeg = 0 To ndes.Count - 1

            nde = ndes(iTrackSeg)
            ndesTrkPt = nde.SelectNodes("*[contains(name(), 'trkpt')]")
            strStart = FormatDateTime(ndesTrkPt, FirstLast.First)
            strEnd = FormatDateTime(ndesTrkPt, FirstLast.Last)

            'We only consider track segments where there is more than one point and which
            'have time elements associated with them.
            If strStart.Length > 0 And strEnd.Length > 0 And ndesTrkPt.Count > 1 Then
                iValidTrack += 1
                str(0) = iValidTrack
                str(1) = strStart
                str(2) = strEnd
                str(3) = ndesTrkPt.Count.ToString
                str(4) = iTrackSeg
                li = New ListViewItem(str)
                lvTracks.Items.Add(li)
            End If
        Next

        If lvTracks.Items.Count = 1 Then

            ndeTrkSeg = ndes(CInt(lvTracks.Items(0).SubItems(4).Text))
            Me.Close()
        ElseIf lvTracks.Items.Count = 0 Then
            ndeTrkSeg = Nothing
            Me.Close()
        Else
            'Do nothing (i.e. show this dialog)
        End If
    End Sub

    Private Function FormatDateTime(ByVal ndesTrkPt As XmlNodeList, ByVal fl As FirstLast) As String

        Dim iPnt As Integer
        Dim strTime As String = ""
        Dim ndeTrkPt As XmlNode
        Dim ndeTime As XmlNode
        Dim dtmThis As DateTime

        If fl = FirstLast.First Then
            iPnt = 0
        Else
            iPnt = ndesTrkPt.Count - 1
        End If

        If ndesTrkPt.Count > 0 Then
            ndeTrkPt = ndesTrkPt(iPnt)
            If Not ndeTrkPt Is Nothing Then
                ndeTime = ndeTrkPt.SelectSingleNode("*[contains(name(), 'time')]")
                If Not ndeTime Is Nothing Then
                    strTime = ndeTime.InnerText
                    dtmThis = DateTime.Parse(strTime)
                End If
            End If
        End If

        If strTime = "" Then
            Return ""
        Else
            Return UTC2LocalTime(dtmThis, RecDateTimeType.RecDate) & " " & UTC2LocalTime(dtmThis, RecDateTimeType.RecTime)
        End If
    End Function

    Private Function UTC2LocalTime(ByVal utcDateTime As DateTime, ByVal dtType As RecDateTimeType) As String

        Dim localDateTime As DateTime = utcDateTime.ToLocalTime()

        If dtType = RecDateTimeType.RecDate Then

            Return (Format(localDateTime, "dd/MM/yyyy"))
        Else
            Return (Format(localDateTime, "HH:mm:ss"))
        End If
    End Function

    Private Sub butOkay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butOkay.Click

        If lvTracks.SelectedItems.Count = 0 Then

            MessageBox.Show("First select a track segment to use.")
        Else
            ndeTrkSeg = ndes(CInt(lvTracks.SelectedItems(0).SubItems(4).Text))
            Me.Close()
        End If
    End Sub

    Private Sub butCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butCancel.Click

        ndeTrkSeg = Nothing
        Me.Close()
    End Sub

    Private Sub frmSelectTrack_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        InitialiseList()

        lvTracks.Columns(0).Width = 70
        lvTracks.Columns(1).Width = 150
        lvTracks.Columns(2).Width = 150
    End Sub
End Class

