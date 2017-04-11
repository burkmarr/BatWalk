Imports System.Math
Imports System.IO

Public Class clsKML

    Shared Sub StyleKML(ByVal srKML As StreamWriter, ByVal colour As Color, ByVal styleName As String)

        srKML.WriteLine("<Style id=""" & styleName & """>")
        srKML.WriteLine("<LabelStyle>")
        srKML.WriteLine("<scale>0</scale>")
        srKML.WriteLine("</LabelStyle>")
        srKML.WriteLine("<BalloonStyle>")
        srKML.WriteLine("<text><u><b>$[name]</b></u><br />$[description]</text>")
        srKML.WriteLine("</BalloonStyle>")

        srKML.WriteLine("<IconStyle>")
        srKML.WriteLine("<color>" & GetGEColour(colour) & "</color>")
        srKML.WriteLine("<scale>0.8</scale>")
        srKML.WriteLine("<Icon>")
        srKML.WriteLine("<href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>")
        srKML.WriteLine("</Icon>")
        srKML.WriteLine("</IconStyle>")

        srKML.WriteLine("<LineStyle>")
        srKML.WriteLine("<color>" & GetGEColour(Color.Black) & "</color>")
        srKML.WriteLine("<width>1</width>")
        srKML.WriteLine("</LineStyle>")
        srKML.WriteLine("<PolyStyle>")
        srKML.WriteLine("<color>" & GetGEColour(colour) & "</color>")
        srKML.WriteLine("<outline>1</outline>")
        srKML.WriteLine("</PolyStyle>")

        srKML.WriteLine("</Style>")
    End Sub

    Shared Function GetGEColour(ByVal colour As Color) As String

        Dim iRed As Integer = colour.R
        Dim iGreen As Integer = colour.G
        Dim iBlue As Integer = colour.B

        'Return "b0" & Hex(iBlue).PadLeft(2, "0") & Hex(iGreen).PadLeft(2, "0") & Hex(iRed).PadLeft(2, "0")
        Return "ff" & Hex(iBlue).PadLeft(2, "0") & Hex(iGreen).PadLeft(2, "0") & Hex(iRed).PadLeft(2, "0")
    End Function

    Shared Sub RenderGR(ByVal srKML As StreamWriter, ByVal sGR As String)

        Dim objGridRef As clsGridRef = New clsGridRef
        objGridRef.MakePrefixArrays()
        objGridRef.GridRef = sGR
        objGridRef.sErrorMessage = ""
        objGridRef.ParseGridRef(True)
        objGridRef.ParseInput(False)

        srKML.WriteLine("<Placemark>")
        srKML.WriteLine("<name>" & sGR & "</name>")


        srKML.WriteLine("<styleUrl>#" & objGridRef.sRefType & "</styleUrl>")
    
        srKML.WriteLine("<visibility>1</visibility>")
        srKML.WriteLine("<open>0</open>")
        srKML.WriteLine("<Polygon>")
        srKML.WriteLine("<tessellate>1</tessellate>")
        srKML.WriteLine("<altitudeMode>clampToGround</altitudeMode>")
        srKML.WriteLine("<outerBoundaryIs>")
        srKML.WriteLine("<LinearRing>")
        srKML.WriteLine("<coordinates>")
        srKML.WriteLine(SquareKML(objGridRef.East, objGridRef.North, objGridRef.fWidth))
        srKML.WriteLine("</coordinates>")
        srKML.WriteLine("</LinearRing>")
        srKML.WriteLine("</outerBoundaryIs>")
        srKML.WriteLine("</Polygon>")
        srKML.WriteLine("</Placemark>")
    End Sub

    Shared Sub LookAtKML(ByVal srKML As StreamWriter, ByVal sGR As String)

        Dim objGridRef As clsGridRef = New clsGridRef
        objGridRef.MakePrefixArrays()
        objGridRef.GridRef = sGR
        objGridRef.ParseGridRef(True)
        objGridRef.ParseInput(False)
        Dim strWidthInMetres As String

        Select Case objGridRef.sRefType

            Case "100km"
                strWidthInMetres = "120000"
            Case "hectad"
                strWidthInMetres = "20000"
            Case "tetrad"
                strWidthInMetres = "8000"
            Case "monad"
                strWidthInMetres = "5000"
            Case "6fig"
                strWidthInMetres = "800"
            Case "8fig"
                strWidthInMetres = "300"
            Case "10fig"
                strWidthInMetres = "50"
            Case Else
                strWidthInMetres = ""
        End Select

        srKML.WriteLine("<LookAt>")
        srKML.WriteLine("<longitude>" & objGridRef.Easting2LongWGS84(objGridRef.East, objGridRef.North, 0) & "</longitude>")
        srKML.WriteLine("<latitude>" & objGridRef.Northing2LatWGS84(objGridRef.East, objGridRef.North, 0) & "</latitude>")
        srKML.WriteLine("<range>" & strWidthInMetres & "</range>")
        srKML.WriteLine("</LookAt>")
    End Sub


    Shared Function SquareKML(ByVal dblEast As Double, ByVal dblNorth As Double, ByVal dblWidth As Double) As String

        Dim strCoords As String

        Dim objGridRef As clsGridRef = New clsGridRef
        objGridRef.MakePrefixArrays()

        strCoords = objGridRef.Easting2LongWGS84(dblEast - dblWidth, dblNorth - dblWidth, 100).ToString & ","
        strCoords = strCoords & objGridRef.Northing2LatWGS84(dblEast - dblWidth, dblNorth - dblWidth, 100).ToString & ",100 "

        strCoords = strCoords & objGridRef.Easting2LongWGS84(dblEast + dblWidth, dblNorth - dblWidth, 100).ToString & ","
        strCoords = strCoords & objGridRef.Northing2LatWGS84(dblEast + dblWidth, dblNorth - dblWidth, 100).ToString & ",100 "

        strCoords = strCoords & objGridRef.Easting2LongWGS84(dblEast + dblWidth, dblNorth + dblWidth, 100).ToString & ","
        strCoords = strCoords & objGridRef.Northing2LatWGS84(dblEast + dblWidth, dblNorth + dblWidth, 100).ToString & ",100 "

        strCoords = strCoords & objGridRef.Easting2LongWGS84(dblEast - dblWidth, dblNorth + dblWidth, 100).ToString & ","
        strCoords = strCoords & objGridRef.Northing2LatWGS84(dblEast - dblWidth, dblNorth + dblWidth, 100).ToString & ",100 "

        Return strCoords
    End Function
End Class
