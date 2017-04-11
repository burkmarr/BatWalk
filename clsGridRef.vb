Imports System.Math

Public Class clsGridRef

    'This code was produced by Richard Burkmar November 2005.
    'Code taken from elsewhere is accredited in further comments.
    'Please feel free to copy and use this code.
    'No liability is accepted for screw-ups!

    Public sGridRef As String
    Public sErrorMessage As String
    Public gsPrefix(50) As String
    Public giEastingPrefix(50) As Integer
    Public giNorthingPrefix(50) As Integer
    Public miElement As Integer
    Public i As Integer
    Public sRefType As String
    Public bSuffixFound As Boolean
    Public fWidth As Double
    Public msPolyline As String
    Public msPoint As String
    Public zoomLevel As Double

    Public s100 As String = ""
    Public sHectad As String = ""
    Public sQuadrantSuffixes As String = ""
    Public sQuadrant() As String = {"", "", "", "", ""}
    Public sTetrad As String = ""
    Public sMonad As String = ""
    Public s6Fig As String = ""
    Public s8Fig As String = ""
    Public s10Fig As String = ""
    Public sEastingLL As String
    Public sNorthingLL As String
    Public sEastingC As String
    Public sNorthingC As String

    Public East As Double
    Public North As Double

    'Lat/Long conversion constants - these are drawn taken from the first
    'page of the spreadsheet produced by the OS to demonstrate coordinate
    'conversions. This could be found here:
    'http://www.gps.gov.uk/additionalInfo/gpsSpreadsheet.asp
    'in November 2005.

    Dim a As Double
    Dim b As Double
    Dim e0 As Double
    Dim n0 As Double
    Dim f0 As Double
    Dim PHI0 As Double
    Dim LAM0 As Double

    Property GridRef() As String
        Get
            Return sGridRef
        End Get
        Set(ByVal Value As String)

            sGridRef = Value.Replace(" ", "")
        End Set
    End Property

    Sub ErrorMessage(ByVal sError As String)

        If sErrorMessage <> "" Then
            sErrorMessage = sErrorMessage & " "
        End If
        sErrorMessage = sErrorMessage & sError
    End Sub

    Sub MakePrefixArrays()
        a = 6377563.396
        b = 6356256.91
        e0 = 400000
        n0 = -100000
        f0 = 0.9996012717
        PHI0 = 49
        LAM0 = -2
        Call ArrayElements("SV", 0, 0)
        Call ArrayElements("SW", 1, 0)
        Call ArrayElements("SX", 2, 0)
        Call ArrayElements("SY", 3, 0)
        Call ArrayElements("SZ", 4, 0)
        Call ArrayElements("TV", 5, 0)
        Call ArrayElements("SR", 1, 1)
        Call ArrayElements("SS", 2, 1)
        Call ArrayElements("ST", 3, 1)
        Call ArrayElements("SU", 4, 1)
        Call ArrayElements("TQ", 5, 1)
        Call ArrayElements("TR", 6, 1)
        Call ArrayElements("SM", 1, 2)
        Call ArrayElements("SN", 2, 2)
        Call ArrayElements("SO", 3, 2)
        Call ArrayElements("SP", 4, 2)
        Call ArrayElements("TL", 5, 2)
        Call ArrayElements("TM", 6, 2)
        Call ArrayElements("SH", 2, 3)
        Call ArrayElements("SJ", 3, 3)
        Call ArrayElements("SK", 4, 3)
        Call ArrayElements("TF", 5, 3)
        Call ArrayElements("TG", 6, 3)
        Call ArrayElements("SC", 2, 4)
        Call ArrayElements("SD", 3, 4)
        Call ArrayElements("SE", 4, 4)
        Call ArrayElements("TA", 5, 4)
        Call ArrayElements("NW", 1, 5)
        Call ArrayElements("NX", 2, 5)
        Call ArrayElements("NY", 3, 5)
        Call ArrayElements("NZ", 4, 5)
        Call ArrayElements("NR", 1, 6)
        Call ArrayElements("NS", 2, 6)
        Call ArrayElements("NT", 3, 6)
        Call ArrayElements("NU", 4, 6)
        Call ArrayElements("NL", 0, 7)
        Call ArrayElements("NM", 1, 7)
        Call ArrayElements("NN", 2, 7)
        Call ArrayElements("NO", 3, 7)
        'Call ArrayElements("HY", 5, 7)
        'Call ArrayElements("HZ", 6, 7)
        Call ArrayElements("NF", 0, 8)
        Call ArrayElements("NG", 1, 8)
        Call ArrayElements("NH", 2, 8)
        Call ArrayElements("NJ", 3, 8)
        Call ArrayElements("NK", 4, 8)
        'Call ArrayElements("HT", 5, 8)
        'Call ArrayElements("HU", 6, 8)
        Call ArrayElements("NA", 0, 9)
        Call ArrayElements("NB", 1, 9)
        Call ArrayElements("NC", 2, 9)
        Call ArrayElements("ND", 3, 9)
        'Call ArrayElements("HT", 3, 10)
        'Call ArrayElements("HU", 4, 10)
        'Call ArrayElements("HP", 4, 11)
    End Sub

    Sub ArrayElements(ByVal sTilePrefix As String, ByVal iEastingPrefix As Integer, ByVal iNorthingPrefix As Integer)

        miElement = miElement + 1

        gsPrefix(miElement) = sTilePrefix
        giEastingPrefix(miElement) = iEastingPrefix
        giNorthingPrefix(miElement) = iNorthingPrefix
    End Sub
    Function EN26fig(ByVal dblEasting As Double, ByVal dblNorthing As Double) As String

        Dim i As Integer
        Dim strEasting As String = CStr(dblEasting \ 1)
        Dim strNorthing As String = CStr(dblNorthing \ 1)

        If strEasting.Length = 5 Then
            strEasting = "0" & strEasting
        End If
        If strNorthing.Length = 5 Then
            strNorthing = "0" & strNorthing
        End If

        Dim strPrefix As String = ""

        For i = 1 To miElement

            If giEastingPrefix(i) = dblEasting \ 100000 And _
                giNorthingPrefix(i) = dblNorthing \ 100000 Then

                strPrefix = gsPrefix(i)
                Exit For
            End If
        Next

        EN26fig = strPrefix & strEasting.Substring(1, 3) & strNorthing.Substring(1, 3)
    End Function

    Function EN28fig(ByVal dblEasting As Double, ByVal dblNorthing As Double) As String

        Dim i As Integer
        Dim strEasting As String = CStr(dblEasting \ 1)
        Dim strNorthing As String = CStr(dblNorthing \ 1)
        Dim strPrefix As String = ""

        If strEasting.Length = 5 Then
            strEasting = "0" & strEasting
        End If
        If strNorthing.Length = 5 Then
            strNorthing = "0" & strNorthing
        End If

        For i = 1 To miElement

            If giEastingPrefix(i) = dblEasting \ 100000 And _
                giNorthingPrefix(i) = dblNorthing \ 100000 Then

                strPrefix = gsPrefix(i)
                Exit For
            End If
        Next

        EN28fig = strPrefix & strEasting.Substring(1, 4) & strNorthing.Substring(1, 4)
    End Function

    Function EN210fig(ByVal dblEasting As Double, ByVal dblNorthing As Double) As String

        Dim i As Integer
        Dim strEasting As String = CStr(dblEasting \ 1)
        Dim strNorthing As String = CStr(dblNorthing \ 1)
        Dim strPrefix As String = ""

        If strEasting.Length = 5 Then
            strEasting = "0" & strEasting
        End If
        If strNorthing.Length = 5 Then
            strNorthing = "0" & strNorthing
        End If

        For i = 1 To miElement

            If giEastingPrefix(i) = dblEasting \ 100000 And _
                giNorthingPrefix(i) = dblNorthing \ 100000 Then

                strPrefix = gsPrefix(i)
                Exit For
            End If
        Next

        EN210fig = strPrefix & strEasting.Substring(1, 5) & strNorthing.Substring(1, 5)
    End Function

    Function EN2Monad(ByVal dblEasting As Double, ByVal dblNorthing As Double) As String

        Dim i As Integer
        Dim strEasting As String = CStr(dblEasting \ 1)
        Dim strNorthing As String = CStr(dblNorthing \ 1)
        Dim strPrefix As String = ""

        If strEasting.Length = 5 Then
            strEasting = "0" & strEasting
        End If
        If strNorthing.Length = 5 Then
            strNorthing = "0" & strNorthing
        End If

        For i = 1 To miElement

            If giEastingPrefix(i) = dblEasting \ 100000 And _
                giNorthingPrefix(i) = dblNorthing \ 100000 Then

                strPrefix = gsPrefix(i)
                Exit For
            End If
        Next

        EN2Monad = strPrefix & strEasting.Substring(1, 2) & strNorthing.Substring(1, 2)
    End Function

    Function EN2Tetrad(ByVal dblEasting As Double, ByVal dblNorthing As Double) As String

        Dim i As Integer
        Dim strEasting As String = CStr(dblEasting \ 1)
        Dim strNorthing As String = CStr(dblNorthing \ 1)
        Dim strPrefix As String = ""

        If strEasting.Length = 5 Then
            strEasting = "0" & strEasting
        End If
        If strNorthing.Length = 5 Then
            strNorthing = "0" & strNorthing
        End If

        For i = 1 To miElement

            If giEastingPrefix(i) = dblEasting \ 100000 And _
                giNorthingPrefix(i) = dblNorthing \ 100000 Then

                strPrefix = gsPrefix(i)
                Exit For
            End If
        Next

        EN2Tetrad = strPrefix & strEasting.Substring(1, 1) & strNorthing.Substring(1, 1) & Mon2TetSuf(strEasting.Substring(2, 1), strNorthing.Substring(2, 1))
    End Function

    Function EN2Hectad(ByVal dblEasting As Double, ByVal dblNorthing As Double) As String

        Dim i As Integer
        Dim strEasting As String = CStr(dblEasting \ 1)
        Dim strNorthing As String = CStr(dblNorthing \ 1)
        Dim strPrefix As String = ""

        If strEasting.Length = 5 Then
            strEasting = "0" & strEasting
        End If
        If strNorthing.Length = 5 Then
            strNorthing = "0" & strNorthing
        End If

        For i = 1 To miElement

            If giEastingPrefix(i) = dblEasting \ 100000 And _
                giNorthingPrefix(i) = dblNorthing \ 100000 Then

                strPrefix = gsPrefix(i)
                Exit For
            End If
        Next

        EN2Hectad = strPrefix & strEasting.Substring(1, 1) & strNorthing.Substring(1, 1)
    End Function

    Function Prefix2Easting(ByVal sPrefix As String) As String

        Dim i As Integer
        Dim sEasting As String = ""

        For i = 1 To miElement

            If gsPrefix(i) = UCase(sPrefix) Then

                sEasting = CStr(giEastingPrefix(i))
                Exit For
            End If
        Next

        Prefix2Easting = sEasting
    End Function

    Function Prefix2Northing(ByVal sPrefix As String) As String

        Dim i As Integer
        Dim sNorthing As String = ""

        For i = 1 To miElement

            If gsPrefix(i) = UCase(sPrefix) Then

                sNorthing = CStr(giNorthingPrefix(i))
                Exit For
            End If
        Next

        Prefix2Northing = sNorthing
    End Function

    Function TetSuf2QuadSuf(ByVal sTetSuf As String) As String

        TetSuf2QuadSuf = ""

        If InStr(1, "ABFG", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "SW"

        ElseIf InStr(1, "DEIJ", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "NW"

        ElseIf InStr(1, "QRVW", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "SE"

        ElseIf InStr(1, "TUYZ", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "NE"

        ElseIf InStr(1, "LK", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "SW,SE"

        ElseIf InStr(1, "PN", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "NW/NE"

        ElseIf InStr(1, "CH", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "NW/SW"

        ElseIf InStr(1, "SX", UCase(sTetSuf)) Then

            TetSuf2QuadSuf = "NE/SE"

        ElseIf UCase(sTetSuf) = "M" Then

            TetSuf2QuadSuf = "NW/NE/SW/SE"
        End If

    End Function

    Function Mon2QuadSuf(ByVal sE As String, ByVal sN As String) As String

        Dim sNS As String
        Dim sEW As String

        If CInt(sE) < 5 Then
            sEW = "W"
        Else
            sEW = "E"
        End If

        If CInt(sN) < 5 Then
            sNS = "S"
        Else
            sNS = "N"
        End If

        Mon2QuadSuf = sNS & sEW
    End Function

    Function Mon2TetSuf(ByVal sE As String, ByVal sN As String) As String

        Dim intE As Integer
        Dim intN As Integer
        intE = CInt(sE)
        intN = CInt(sN)

        Mon2TetSuf = ""

        If intE < 2 And intN < 2 Then Mon2TetSuf = "A"
        If intE < 2 And intN >= 2 And Int(sN) < 4 Then Mon2TetSuf = "B"
        If intE < 2 And intN >= 4 And Int(sN) < 6 Then Mon2TetSuf = "C"
        If intE < 2 And intN >= 6 And Int(sN) < 8 Then Mon2TetSuf = "D"
        If intE < 2 And intN >= 8 Then Mon2TetSuf = "E"
        If intE < 4 And intE >= 2 And intN < 2 Then Mon2TetSuf = "F"
        If intE < 4 And intE >= 2 And intN >= 2 And intN < 4 Then Mon2TetSuf = "G"
        If intE < 4 And intE >= 2 And intN >= 4 And intN < 6 Then Mon2TetSuf = "H"
        If intE < 4 And intE >= 2 And intN >= 6 And intN < 8 Then Mon2TetSuf = "I"
        If intE < 4 And intE >= 2 And intN >= 8 Then Mon2TetSuf = "J"
        If intE < 6 And intE >= 4 And intN < 2 Then Mon2TetSuf = "K"
        If intE < 6 And intE >= 4 And intN >= 2 And intN < 4 Then Mon2TetSuf = "L"
        If intE < 6 And intE >= 4 And intN >= 4 And intN < 6 Then Mon2TetSuf = "M"
        If intE < 6 And intE >= 4 And intN >= 6 And intN < 8 Then Mon2TetSuf = "N"
        If intE < 6 And intE >= 4 And intN >= 8 Then Mon2TetSuf = "P"
        If intE < 8 And intE >= 6 And intN < 2 Then Mon2TetSuf = "Q"
        If intE < 8 And intE >= 6 And intN >= 2 And intN < 4 Then Mon2TetSuf = "R"
        If intE < 8 And intE >= 6 And intN >= 4 And intN < 6 Then Mon2TetSuf = "S"
        If intE < 8 And intE >= 6 And intN >= 6 And intN < 8 Then Mon2TetSuf = "T"
        If intE < 8 And intE >= 6 And intN >= 8 Then Mon2TetSuf = "U"
        If intE >= 8 And intN < 2 Then Mon2TetSuf = "V"
        If intE >= 8 And intN >= 2 And intN < 4 Then Mon2TetSuf = "W"
        If intE >= 8 And intN >= 4 And intN < 6 Then Mon2TetSuf = "X"
        If intE >= 8 And intN >= 6 And intN < 8 Then Mon2TetSuf = "Y"
        If intE >= 8 And intN >= 8 Then Mon2TetSuf = "Z"
    End Function

    Function QuadSuf2DX(ByVal sSuffix As String, ByVal bCentre As String) As String

        If UCase(sSuffix) = "NW" Or UCase(sSuffix) = "SW" Then

            If bCentre Then

                QuadSuf2DX = "2500"
            Else
                QuadSuf2DX = "0000"
            End If
        Else
            If bCentre Then

                QuadSuf2DX = "7500"
            Else
                QuadSuf2DX = "5000"
            End If
        End If
    End Function

    Function QuadSuf2DY(ByVal sSuffix As String, ByVal bCentre As String) As String

        If UCase(sSuffix) = "SW" Or UCase(sSuffix) = "SE" Then

            If bCentre Then

                QuadSuf2DY = "2500"
            Else
                QuadSuf2DY = "0000"
            End If
        Else
            If bCentre Then

                QuadSuf2DY = "7500"
            Else
                QuadSuf2DY = "5000"
            End If
        End If
    End Function

    Function TetSuf2DX(ByVal sSuffix As String, ByVal bCentre As String) As String

        If InStr(1, "ABCDE", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DX = "1000"
            Else
                TetSuf2DX = "0000"
            End If
        ElseIf InStr(1, "FGHIJ", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DX = "3000"
            Else
                TetSuf2DX = "2000"
            End If
        ElseIf InStr(1, "KLMNP", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DX = "5000"
            Else
                TetSuf2DX = "4000"
            End If
        ElseIf InStr(1, "QRSTU", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DX = "7000"
            Else
                TetSuf2DX = "6000"
            End If
        Else
            If bCentre Then

                TetSuf2DX = "9000"
            Else
                TetSuf2DX = "8000"
            End If
        End If

    End Function

    Function TetSuf2DY(ByVal sSuffix As String, ByVal bCentre As String) As String

        If InStr(1, "AFKQV", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DY = "1000"
            Else
                TetSuf2DY = "0000"
            End If
        ElseIf InStr(1, "BGLRW", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DY = "3000"
            Else
                TetSuf2DY = "2000"
            End If
        ElseIf InStr(1, "CHMSX", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DY = "5000"
            Else
                TetSuf2DY = "4000"
            End If
        ElseIf InStr(1, "DINTY", UCase(sSuffix)) Then

            If bCentre Then

                TetSuf2DY = "7000"
            Else
                TetSuf2DY = "6000"
            End If
        Else
            If bCentre Then

                TetSuf2DY = "9000"
            Else
                TetSuf2DY = "8000"
            End If
        End If

    End Function

    Function TestNumeric(ByVal sString As String) As Boolean

        Dim i As Integer

        For i = 1 To Len(sString)

            If Not InStr(1, "0123456789", Mid(sString, i, 1)) > 0 Then
                TestNumeric = False
                Exit Function
            End If
        Next
        TestNumeric = True
    End Function

    Function TestQuadrantSuffix(ByVal sString As String) As Boolean

        If Not UCase(sString) = "SW" And Not UCase(sString) = "SE" And Not UCase(sString) = "NW" And Not UCase(sString) = "NE" Then

            TestQuadrantSuffix = False
            Exit Function
        End If

        TestQuadrantSuffix = True
    End Function

    Function TestTetradSuffix(ByVal sString As String) As Boolean

        If Not InStr(1, "ABCDEFGHIJKLMNPQRSTUVWXYZ", UCase(sString)) > 0 Then
            TestTetradSuffix = False
            Exit Function
        End If

        TestTetradSuffix = True
    End Function

    Sub ParseGridRef(ByVal bMin As Boolean)

        Dim i As Integer

        If Len(sGridRef) < 2 Then

            ErrorMessage("The grid reference must be at least two characters long.")
        ElseIf Len(sGridRef) = 2 Then
            sRefType = "100km"
        End If

        bSuffixFound = False
        For i = 1 To miElement

            If gsPrefix(i) = UCase(Left(sGridRef, 2)) Then

                bSuffixFound = True
                Exit For
            End If
        Next
        If Not bSuffixFound Then

            ErrorMessage(UCase(Left(sGridRef, 2)) & " is not a valid suffix for a grid reference.")
        End If

        If Len(sGridRef) = 3 Then

            ErrorMessage("A grid reference cannot be 3 characters long.")
        End If
        If Len(sGridRef) = 7 Then

            ErrorMessage("A grid reference cannot be 7 characters long.")
        End If
        If Len(sGridRef) = 9 Then

            ErrorMessage("A grid reference cannot be 9 characters long.")
        End If
        If Len(sGridRef) = 11 Then

            ErrorMessage("A grid reference cannot be 11 characters long.")
        End If
        If Len(sGridRef) > 12 Then

            ErrorMessage("The grid reference must not be more than 12 characters long.")
        End If

        If Len(sGridRef) = 4 Then
            sRefType = "hectad"
            If Not TestNumeric(Mid(sGridRef, 3, 2)) Then
                ErrorMessage("Invalid hectad reference: characters 3 and 4 must be numeric.")
            End If
        End If

        If Len(sGridRef) = 5 Then
            sRefType = "tetrad"
            If Not TestNumeric(Mid(sGridRef, 3, 2)) Then

                ErrorMessage("Invalid tetrad reference: characters 3 and 4 must be numeric.")
            End If
            If Not TestTetradSuffix(Right(sGridRef, 1)) Then

                ErrorMessage("Invalid tetrad reference: last character must be a letter (but not o).")
            End If
        End If

        If Len(sGridRef) = 6 Then

            If TestNumeric(Mid(sGridRef, 3, 4)) Then
                sRefType = "monad"
            Else
                sRefType = "quadrant"
                If Not TestNumeric(Mid(sGridRef, 3, 2)) Then

                    ErrorMessage("Invalid monad or quadrant reference: characters 3 and 4 must be numeric.")

                ElseIf Not TestQuadrantSuffix(Right(sGridRef, 2)) Then

                    ErrorMessage("Invalid quadrant reference: last two characters must be SW, SE, NW or NE.")
                Else
                    If bMin Then

                        'ErrorMessage("You've specified a quadrant reference, but for the atlas the grid reference must resolve at least to tetrad level.")
                    End If
                End If
            End If
        End If

        If Len(sGridRef) = 8 Then

            sRefType = "6fig"
            If Not TestNumeric(Mid(sGridRef, 3, 6)) Then

                ErrorMessage("Invalid 6 figure grid reference: characters 3-8 must be numeric.")
            End If
        End If

        If Len(sGridRef) = 10 Then

            sRefType = "8fig"
            If Not TestNumeric(Mid(sGridRef, 3, 8)) Then

                ErrorMessage("Invalid 8 figure grid reference: characters 3-10 must be numeric.")
            End If
        End If

        If Len(sGridRef) = 12 Then

            sRefType = "10fig"
            If Not TestNumeric(Mid(sGridRef, 3, 10)) Then

                ErrorMessage("Invalid 10 figure grid reference: characters 3-12 must be numeric.")
            End If
        End If

        If bMin And Len(sGridRef) < 5 Then

            'ErrorMessage("For the atlas, the grid reference must resolve at least to tetrad level (i.e. must be five characters or more).")
        End If
    End Sub

    Function E_N_to_Lat(ByVal East As Double, ByVal North As Double, ByVal a As Double, ByVal b As Double, ByVal e0 As Double, ByVal n0 As Double, ByVal f0 As Double, ByVal PHI0 As Double, ByVal LAM0 As Double) As Double

        'This function  taken from the first
        'the spreadsheet produced by the OS to demonstrate coordinate
        'conversions. This  could be found here:
        'http://www.gps.gov.uk/additionalInfo/gpsSpreadsheet.asp
        'in November 2005.

        'Un-project Transverse Mercator eastings and northings back to latitude.
        'Input: - _
        'eastings (East) and northings (North) in meters; _
        ' ellipsoid axis dimensions (a & b) in meters; _
        ' eastings (e0) and northings (n0) of false origin in meters; _
        ' central meridian scale factor (f0) and _
        ' latitude (PHI0) and longitude (LAM0) of false origin in decimal degrees.

        'REQUIRES THE "Marc" AND "InitialLat" FUNCTIONS

        Dim Pi As Double
        Dim RadPHI0 As Double
        Dim RadLAM0 As Double
        Dim af0 As Double
        Dim bf0 As Double
        Dim e2 As Double
        Dim n As Double
        Dim Et As Double
        Dim PHId As Double
        Dim nu As Double
        Dim rho As Double
        Dim eta2 As Double
        Dim VII As Double
        Dim VIII As Double
        Dim IX As Double

        'Convert angle measures to radians
        Pi = 3.14159265358979
        RadPHI0 = PHI0 * (Pi / 180)
        RadLAM0 = LAM0 * (Pi / 180)

        'Compute af0, bf0, e squared (e2), n and Et
        af0 = a * f0
        bf0 = b * f0
        e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
        n = (af0 - bf0) / (af0 + bf0)
        Et = East - e0

        'Compute initial value for latitude (PHI) in radians
        PHId = InitialLat(North, n0, af0, RadPHI0, n, bf0)

        'Compute nu, rho and eta2 using value for PHId
        nu = af0 / (Sqrt(1 - (e2 * ((Sin(PHId)) ^ 2))))
        rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))
        eta2 = (nu / rho) - 1

        'Compute Latitude
        VII = (Tan(PHId)) / (2 * rho * nu)
        VIII = ((Tan(PHId)) / (24 * rho * (nu ^ 3))) * (5 + (3 * ((Tan(PHId)) ^ 2)) + eta2 - (9 * eta2 * ((Tan(PHId)) ^ 2)))
        IX = ((Tan(PHId)) / (720 * rho * (nu ^ 5))) * (61 + (90 * ((Tan(PHId)) ^ 2)) + (45 * ((Tan(PHId)) ^ 4)))

        E_N_to_Lat = (180 / Pi) * (PHId - ((Et ^ 2) * VII) + ((Et ^ 4) * VIII) - ((Et ^ 6) * IX))

    End Function

    Function Marc(ByVal bf0 As Double, ByVal n As Double, ByVal PHI0 As Double, ByVal PHI As Double) As Double
        'This function  taken from the first
        'the spreadsheet produced by the OS to demonstrate coordinate
        'conversions. This  could be found here:
        'http://www.gps.gov.uk/additionalInfo/gpsSpreadsheet.asp
        'in November 2005.

        'Compute meridional arc.
        'Input: - _
        'ellipsoid semi major axis multiplied by central meridian scale factor (bf0) in meters; _
        'n (computed from a, b and f0); _
        'lat of false origin (PHI0) and initial or final latitude of point (PHI) IN RADIANS.

        'THIS FUNCTION IS CALLED BY THE - "Lat_Long_to_North" and "InitialLat" FUNCTIONS
        'THIS FUNCTION IS ALSO USED ON IT'S OWN IN THE "Projection and Transformation Calculations.xls" SPREADSHEET

        Marc = bf0 * (((1 + n + ((5 / 4) * (n ^ 2)) + ((5 / 4) * (n ^ 3))) * (PHI - PHI0)) _
        - (((3 * n) + (3 * (n ^ 2)) + ((21 / 8) * (n ^ 3))) * (Sin(PHI - PHI0)) * (Cos(PHI + PHI0))) _
        + ((((15 / 8) * (n ^ 2)) + ((15 / 8) * (n ^ 3))) * (Sin(2 * (PHI - PHI0))) * (Cos(2 * (PHI + PHI0)))) _
        - (((35 / 24) * (n ^ 3)) * (Sin(3 * (PHI - PHI0))) * (Cos(3 * (PHI + PHI0)))))

    End Function

    Function InitialLat(ByVal North As Double, ByVal n0 As Double, ByVal afo As Double, ByVal PHI0 As Double, ByVal n As Double, ByVal bfo As Double) As Double
        'This function  taken from the first
        'the spreadsheet produced by the OS to demonstrate coordinate
        'conversions. This  could be found here:
        'http://www.gps.gov.uk/additionalInfo/gpsSpreadsheet.asp
        'in November 2005.

        'Compute initial value for Latitude (PHI) IN RADIANS.
        'Input: - _
        ' northing of point (North) and northing of false origin (n0) in meters; _
        ' semi major axis multiplied by central meridian scale factor (af0) in meters; _
        ' latitude of false origin (PHI0) IN RADIANS; _
        ' n (computed from a, b and f0) and _
        ' ellipsoid semi major axis multiplied by central meridian scale factor (bf0) in meters.

        'REQUIRES THE "Marc" FUNCTION
        'THIS FUNCTION IS CALLED BY THE "E_N_to_Lat", "E_N_to_Long" and "E_N_to_C" FUNCTIONS
        'THIS FUNCTION IS ALSO USED ON IT'S OWN IN THE  "Projection and Transformation Calculations.xls" SPREADSHEET

        Dim PHI1 As Double
        Dim M As Double
        Dim PHI2 As Double

        'First PHI value (PHI1)
        PHI1 = ((North - n0) / afo) + PHI0

        'Calculate M
        M = Marc(bfo, n, PHI0, PHI1)

        'Calculate new PHI value (PHI2)
        PHI2 = ((North - n0 - M) / afo) + PHI1

        'Iterate to get final value for InitialLat
        Do While Abs(North - n0 - M) > 0.00001
            PHI2 = ((North - n0 - M) / afo) + PHI1
            M = Marc(bfo, n, PHI0, PHI2)
            PHI1 = PHI2
        Loop

        InitialLat = PHI2

    End Function

    Function E_N_to_Long(ByVal East As Double, ByVal North As Double, ByVal a As Double, ByVal b As Double, ByVal e0 As Double, ByVal n0 As Double, ByVal f0 As Double, ByVal PHI0 As Double, ByVal LAM0 As Double) As Double
        'This function  taken from the first
        'the spreadsheet produced by the OS to demonstrate coordinate
        'conversions. This  could be found here:
        'http://www.gps.gov.uk/additionalInfo/gpsSpreadsheet.asp
        'in November 2005.

        'Un-project Transverse Mercator eastings and northings back to longitude.
        'Input: - _
        ' eastings (East) and northings (North) in meters; _
        ' ellipsoid axis dimensions (a & b) in meters; _
        ' eastings (e0) and northings (n0) of false origin in meters; _
        ' central meridian scale factor (f0) and _
        ' latitude (PHI0) and longitude (LAM0) of false origin in decimal degrees.

        'REQUIRES THE "Marc" AND "InitialLat" FUNCTIONS

        Dim Pi As Double
        Dim RadPHI0 As Double
        Dim RadLAM0 As Double
        Dim af0 As Double
        Dim bf0 As Double
        Dim e2 As Double
        Dim n As Double
        Dim Et As Double
        Dim PHId As Double
        Dim nu As Double
        Dim rho As Double
        Dim eta2 As Double
        Dim X As Double
        Dim XI As Double
        Dim XII As Double
        Dim XIIA As Double

        'Convert angle measures to radians
        Pi = 3.14159265358979
        RadPHI0 = PHI0 * (Pi / 180)
        RadLAM0 = LAM0 * (Pi / 180)

        'Compute af0, bf0, e squared (e2), n and Et
        af0 = a * f0
        bf0 = b * f0
        e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
        n = (af0 - bf0) / (af0 + bf0)
        Et = East - e0

        'Compute initial value for latitude (PHI) in radians
        PHId = InitialLat(North, n0, af0, RadPHI0, n, bf0)

        'Compute nu, rho and eta2 using value for PHId
        nu = af0 / (Sqrt(1 - (e2 * ((Sin(PHId)) ^ 2))))
        rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(PHId)) ^ 2))
        eta2 = (nu / rho) - 1

        'Compute Longitude
        X = ((Cos(PHId)) ^ -1) / nu
        XI = (((Cos(PHId)) ^ -1) / (6 * (nu ^ 3))) * ((nu / rho) + (2 * ((Tan(PHId)) ^ 2)))
        XII = (((Cos(PHId)) ^ -1) / (120 * (nu ^ 5))) * (5 + (28 * ((Tan(PHId)) ^ 2)) + (24 * ((Tan(PHId)) ^ 4)))
        XIIA = (((Cos(PHId)) ^ -1) / (5040 * (nu ^ 7))) * (61 + (662 * ((Tan(PHId)) ^ 2)) + (1320 * ((Tan(PHId)) ^ 4)) + (720 * ((Tan(PHId)) ^ 6)))

        E_N_to_Long = (180 / Pi) * (RadLAM0 + (Et * X) - ((Et ^ 3) * XI) + ((Et ^ 5) * XII) - ((Et ^ 7) * XIIA))

    End Function

    Function Lat_Long_to_East(ByVal PHI As Double, ByVal LAM As Double, ByVal a As Double, ByVal b As Double, ByVal e0 As Double, ByVal f0 As Double, ByVal PHI0 As Double, ByVal LAM0 As Double) As Double
        'Project Latitude and longitude to Transverse Mercator eastings.
        'Input: -
        'Latitude (PHI) and Longitude (LAM) in decimal degrees;
        'ellipsoid axis dimensions (a & b) in meters;
        'eastings of false origin (e0) in meters;
        'central meridian scale factor (f0);
        'latitude (PHI0) and longitude (LAM0) of false origin in decimal degrees.

        Dim Pi As Double
        Dim RadPHI0 As Double
        Dim RadLAM0 As Double
        Dim RadPHI As Double
        Dim RadLAM As Double
        Dim af0 As Double
        Dim bf0 As Double
        Dim e2 As Double
        Dim n As Double
        Dim nu As Double
        Dim rho As Double
        Dim eta2 As Double
        Dim p As Double
        Dim IV As Double
        Dim V As Double
        Dim VI As Double

        'Convert angle measures to radians
        Pi = 3.14159265358979
        RadPHI = PHI * (Pi / 180)
        RadLAM = LAM * (Pi / 180)
        RadPHI0 = PHI0 * (Pi / 180)
        RadLAM0 = LAM0 * (Pi / 180)

        af0 = a * f0
        bf0 = b * f0
        e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
        n = (af0 - bf0) / (af0 + bf0)
        nu = af0 / (Sqrt(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
        rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(RadPHI)) ^ 2))
        eta2 = (nu / rho) - 1
        p = RadLAM - RadLAM0

        IV = nu * (Cos(RadPHI))
        V = (nu / 6) * ((Cos(RadPHI)) ^ 3) * ((nu / rho) - ((Tan(RadPHI) ^ 2)))
        VI = (nu / 120) * ((Cos(RadPHI)) ^ 5) * (5 - (18 * ((Tan(RadPHI)) ^ 2)) + ((Tan(RadPHI)) ^ 4) + (14 * eta2) - (58 * ((Tan(RadPHI)) ^ 2) * eta2))

        Lat_Long_to_East = e0 + (p * IV) + ((p ^ 3) * V) + ((p ^ 5) * VI)

    End Function

    Function Lat_Long_to_North(ByVal PHI As Double, ByVal LAM As Double, ByVal a As Double, ByVal b As Double, ByVal e0 As Double, ByVal n0 As Double, ByVal f0 As Double, ByVal PHI0 As Double, ByVal LAM0 As Double) As Double
        'Project Latitude and longitude to Transverse Mercator northings
        'Input: - Latitude (PHI) and Longitude (LAM) in decimal degrees;
        'ellipsoid axis dimensions (a & b) in meters;
        'eastings (e0) and northings (n0) of false origin in meters;
        'central meridian scale factor (f0);
        'latitude (PHI0) and longitude (LAM0) of false origin in decimal degrees.

        'REQUIRES THE "Marc" FUNCTION

        Dim Pi As Double
        Dim RadPHI0 As Double
        Dim RadLAM0 As Double
        Dim RadPHI As Double
        Dim RadLAM As Double
        Dim af0 As Double
        Dim bf0 As Double
        Dim e2 As Double
        Dim n As Double
        Dim nu As Double
        Dim rho As Double
        Dim eta2 As Double
        Dim p As Double
        Dim M As Double
        Dim I As Double
        Dim II As Double
        Dim III As Double
        Dim IIIa As Double

        'Convert angle measures to radians
        Pi = 3.14159265358979
        RadPHI = PHI * (Pi / 180)
        RadLAM = LAM * (Pi / 180)
        RadPHI0 = PHI0 * (Pi / 180)
        RadLAM0 = LAM0 * (Pi / 180)

        af0 = a * f0
        bf0 = b * f0
        e2 = ((af0 ^ 2) - (bf0 ^ 2)) / (af0 ^ 2)
        n = (af0 - bf0) / (af0 + bf0)
        nu = af0 / (Sqrt(1 - (e2 * ((Sin(RadPHI)) ^ 2))))
        rho = (nu * (1 - e2)) / (1 - (e2 * (Sin(RadPHI)) ^ 2))
        eta2 = (nu / rho) - 1
        p = RadLAM - RadLAM0
        M = Marc(bf0, n, RadPHI0, RadPHI)

        I = M + n0
        II = (nu / 2) * (Sin(RadPHI)) * (Cos(RadPHI))
        III = ((nu / 24) * (Sin(RadPHI)) * ((Cos(RadPHI)) ^ 3)) * (5 - ((Tan(RadPHI)) ^ 2) + (9 * eta2))
        IIIa = ((nu / 720) * (Sin(RadPHI)) * ((Cos(RadPHI)) ^ 5)) * (61 - (58 * ((Tan(RadPHI)) ^ 2)) + ((Tan(RadPHI)) ^ 4))

        Lat_Long_to_North = I + ((p ^ 2) * II) + ((p ^ 4) * III) + ((p ^ 6) * IIIa)

    End Function

    Function Lat_Long_H_to_X(ByVal PHI As Double, ByVal LAM As Double, ByVal H As Double, ByVal a As Double, ByVal b As Double) As Double
        'Convert geodetic coords lat (PHI), long (LAM) and height (H) to cartesian X coordinate.
        'Input: - 
        'Latitude (PHI)& Longitude (LAM) both in decimal degrees; 
        'Ellipsoidal height (H) and ellipsoid axis dimensions (a & b) all in meters.


        'Convert angle measures to radians
        Dim Pi As Double = 3.14159265358979
        Dim RadPHI As Double = PHI * (Pi / 180)
        Dim RadLAM As Double = LAM * (Pi / 180)

        'Compute eccentricity squared and nu
        Dim e2 As Double = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
        Dim V As Double = a / (Sqrt(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

        'Compute X
        Lat_Long_H_to_X = (V + H) * (Cos(RadPHI)) * (Cos(RadLAM))

    End Function

    Function Lat_Long_H_to_Y(ByVal PHI As Double, ByVal LAM As Double, ByVal H As Double, ByVal a As Double, ByVal b As Double) As Double
        'Convert geodetic coords lat (PHI), long (LAM) and height (H) to cartesian Y coordinate.
        'Input: -
        'Latitude (PHI)& Longitude (LAM) both in decimal degrees;
        'Ellipsoidal height (H) and ellipsoid axis dimensions (a & b) all in meters.

        'Convert angle measures to radians
        Dim PI As Double = 3.14159265358979
        Dim RadPHI As Double = PHI * (PI / 180)
        Dim RadLAM As Double = LAM * (PI / 180)

        'Compute eccentricity squared and nu
        Dim e2 As Double = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
        Dim V As Double = a / (Sqrt(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

        'Compute Y
        Lat_Long_H_to_Y = (V + H) * (Cos(RadPHI)) * (Sin(RadLAM))

    End Function

    Function Lat_H_to_Z(ByVal PHI As Double, ByVal H As Double, ByVal a As Double, ByVal b As Double) As Double
        'Convert geodetic coord components latitude (PHI) and height (H) to cartesian Z coordinate.
        'Input: -
        'Latitude (PHI) decimal degrees;
        'Ellipsoidal height (H) and ellipsoid axis dimensions (a & b) all in meters.

        'Convert angle measures to radians
        Dim PI As Double = 3.14159265358979
        Dim RadPHI As Double = PHI * (PI / 180)

        'Compute eccentricity squared and nu
        Dim e2 As Double = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
        Dim V As Double = a / (Sqrt(1 - (e2 * ((Sin(RadPHI)) ^ 2))))

        'Compute Z
        Lat_H_to_Z = ((V * (1 - e2)) + H) * (Sin(RadPHI))

    End Function

    Function Helmert_X(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal DX As Double, ByVal Y_Rot As Double, ByVal Z_Rot As Double, ByVal s As Double) As Double
        'Computed Helmert transformed X coordinate.
        'Input: -
        'cartesian XYZ coords (X,Y,Z), X translation (DX) all in meters ;
        'Y and Z rotations in seconds of arc (Y_Rot, Z_Rot) and scale in ppm (s).

        'Convert rotations to radians and ppm scale to a factor
        Dim PI As Double = 3.14159265358979
        Dim sfactor As Double = s * 0.000001
        Dim RadY_Rot As Double = (Y_Rot / 3600) * (PI / 180)
        Dim RadZ_Rot As Double = (Z_Rot / 3600) * (PI / 180)

        'Compute transformed X coord
        Helmert_X = X + (X * sfactor) - (Y * RadZ_Rot) + (Z * RadY_Rot) + DX

    End Function

    Function Helmert_Y(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal DY As Double, ByVal X_Rot As Double, ByVal Z_Rot As Double, ByVal s As Double) As Double
        'Computed Helmert transformed Y coordinate.
        'Input: -
        'cartesian XYZ coords (X,Y,Z), Y translation (DY) all in meters ;
        'X and Z rotations in seconds of arc (X_Rot, Z_Rot) and scale in ppm (s).

        'Convert rotations to radians and ppm scale to a factor
        Dim PI As Double = 3.14159265358979
        Dim sfactor As Double = s * 0.000001
        Dim RadX_Rot As Double = (X_Rot / 3600) * (PI / 180)
        Dim RadZ_Rot As Double = (Z_Rot / 3600) * (PI / 180)

        'Compute transformed Y coord
        Helmert_Y = (X * RadZ_Rot) + Y + (Y * sfactor) - (Z * RadX_Rot) + DY

    End Function

    Function Helmert_Z(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal DZ As Double, ByVal X_Rot As Double, ByVal Y_Rot As Double, ByVal s As Double) As Double
        'Computed Helmert transformed Z coordinate.
        'Input: -
        'cartesian XYZ coords (X,Y,Z), Z translation (DZ) all in meters ;
        'X and Y rotations in seconds of arc (X_Rot, Y_Rot) and scale in ppm (s).

        'Convert rotations to radians and ppm scale to a factor
        Dim PI As Double = 3.14159265358979
        Dim sfactor As Double = s * 0.000001
        Dim RadX_Rot As Double = (X_Rot / 3600) * (PI / 180)
        Dim RadY_Rot As Double = (Y_Rot / 3600) * (PI / 180)

        'Compute transformed Z coord
        Helmert_Z = (-1 * X * RadY_Rot) + (Y * RadX_Rot) + Z + (Z * sfactor) + DZ

    End Function

    Function XYZ_to_Lat(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal a As Double, ByVal b As Double) As Double

        'Convert XYZ to Latitude (PHI) in Dec Degrees.
        'Input: -
        'XYZ cartesian coords (X,Y,Z) and ellipsoid axis dimensions (a & b), all in meters.

        'THIS FUNCTION REQUIRES THE "Iterate_XYZ_to_Lat" FUNCTION
        'THIS FUNCTION IS CALLED BY THE "XYZ_to_H" FUNCTION

        Dim RootXYSqr As Double = Sqrt((X ^ 2) + (Y ^ 2))
        Dim e2 As Double = ((a ^ 2) - (b ^ 2)) / (a ^ 2)
        Dim PHI1 As Double = Atan(Z / (RootXYSqr * (1 - e2)))

        Dim PHI As Double = Iterate_XYZ_to_Lat(a, e2, PHI1, Z, RootXYSqr)

        Dim PI As Double = 3.14159265358979

        XYZ_to_Lat = PHI * (180 / PI)

    End Function

    Function Iterate_XYZ_to_Lat(ByVal a As Double, ByVal e2 As Double, ByVal PHI1 As Double, ByVal Z As Double, ByVal RootXYSqr As Double) As Double
        'Iteratively computes Latitude (PHI).
        'Input: -
        'ellipsoid semi major axis (a) in meters;
        'eta squared (e2);
        'estimated value for latitude (PHI1) in radians;
        'cartesian Z coordinate (Z) in meters;
        'RootXYSqr computed from X & Y in meters.

        'THIS FUNCTION IS CALLED BY THE "XYZ_to_PHI" FUNCTION

        Dim V As Double = a / (Sqrt(1 - (e2 * ((Sin(PHI1)) ^ 2))))
        Dim PHI2 As Double = Atan((Z + (e2 * V * (Sin(PHI1)))) / RootXYSqr)

        Do While Abs(PHI1 - PHI2) > 0.000000001
            PHI1 = PHI2
            V = a / (Sqrt(1 - (e2 * ((Sin(PHI1)) ^ 2))))
            PHI2 = Atan((Z + (e2 * V * (Sin(PHI1)))) / RootXYSqr)
        Loop

        Iterate_XYZ_to_Lat = PHI2

    End Function

    Function XYZ_to_Long(ByVal X As Double, ByVal Y As Double) As Double
        'Convert XYZ to Longitude (LAM) in Dec Degrees.
        'Input: -
        'X and Y cartesian coords in meters.

        Dim PI As Double = 3.14159265358979
        XYZ_to_Long = (Atan(Y / X)) * (180 / PI)

    End Function


    Function Easting2LongWGS84(ByVal East As Double, ByVal North As Double, ByVal ODNHeight As Double) As Double

        Dim LatOSGB36 As Double
        Dim LongOSGB36 As Double
        Dim CartesianXOSGB36 As Double
        Dim CartesianYOSGB36 As Double
        Dim CartesianZOSGB36 As Double
        Dim CartesianXWGS84 As Double
        Dim CartesianYWGS84 As Double
        Dim CartesianZWGS84 As Double

        Dim DX As Double = 446.448 'translation parallel to X 
        Dim DY As Double = -125.157 'translation parallel to Y 
        Dim DZ As Double = 542.06 'translation parallel to Z 
        Dim s As Double = -20.4894 'scale change 
        Dim X_Rot As Double = 0.1502 'rotation about X 
        Dim Y_Rot As Double = 0.247 'rotation about Y 
        Dim Z_Rot As Double = 0.8421 'rotation about Z 

        LongOSGB36 = E_N_to_Long(East, North, a, b, e0, n0, f0, PHI0, LAM0)
        LatOSGB36 = E_N_to_Lat(East, North, a, b, e0, n0, f0, PHI0, LAM0)

        CartesianXOSGB36 = Lat_Long_H_to_X(LatOSGB36, LongOSGB36, ODNHeight, a, b)
        CartesianYOSGB36 = Lat_Long_H_to_Y(LatOSGB36, LongOSGB36, ODNHeight, a, b)
        CartesianZOSGB36 = Lat_H_to_Z(LatOSGB36, ODNHeight, a, b)

        CartesianXWGS84 = Helmert_X(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, DX, Y_Rot, Z_Rot, s)
        CartesianYWGS84 = Helmert_Y(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, DY, X_Rot, Z_Rot, s)
        CartesianZWGS84 = Helmert_Z(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, DZ, X_Rot, Y_Rot, s)

        Easting2LongWGS84 = XYZ_to_Long(CartesianXWGS84, CartesianYWGS84)

    End Function

    Function Northing2LatWGS84(ByVal East As Double, ByVal North As Double, ByVal ODNHeight As Double) As Double

        Dim LatOSGB36 As Double
        Dim LongOSGB36 As Double
        Dim CartesianXOSGB36 As Double
        Dim CartesianYOSGB36 As Double
        Dim CartesianZOSGB36 As Double
        Dim CartesianXWGS84 As Double
        Dim CartesianYWGS84 As Double
        Dim CartesianZWGS84 As Double

        Dim DX As Double = 446.448 'translation parallel to X 
        Dim DY As Double = -125.157 'translation parallel to Y 
        Dim DZ As Double = 542.06 'translation parallel to Z 
        Dim s As Double = -20.4894 'scale change 
        Dim X_Rot As Double = 0.1502 'rotation about X 
        Dim Y_Rot As Double = 0.247 'rotation about Y 
        Dim Z_Rot As Double = 0.8421 'rotation about Z 

        LongOSGB36 = E_N_to_Long(East, North, a, b, e0, n0, f0, PHI0, LAM0)
        LatOSGB36 = E_N_to_Lat(East, North, a, b, e0, n0, f0, PHI0, LAM0)

        CartesianXOSGB36 = Lat_Long_H_to_X(LatOSGB36, LongOSGB36, ODNHeight, a, b)
        CartesianYOSGB36 = Lat_Long_H_to_Y(LatOSGB36, LongOSGB36, ODNHeight, a, b)
        CartesianZOSGB36 = Lat_H_to_Z(LatOSGB36, ODNHeight, a, b)

        CartesianXWGS84 = Helmert_X(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, DX, Y_Rot, Z_Rot, s)
        CartesianYWGS84 = Helmert_Y(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, DY, X_Rot, Z_Rot, s)
        CartesianZWGS84 = Helmert_Z(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, DZ, X_Rot, Y_Rot, s)

        'Ellipsoid parameters WGS84
        Dim aWGS84 As Double = 6378137.0
        Dim bWGS84 As Double = 6356752.314

        Northing2LatWGS84 = XYZ_to_Lat(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, aWGS84, bWGS84)

        'Dim rtmp As Double = Northing2LatWGS84
    End Function


    Public Function WGS84Long2OSGGB36(ByVal LatWGS84 As Double, ByVal LonWGS84 As Double, ByVal ODNHeight As Double) As Double

        Dim CartesianXOSGB36 As Double
        Dim CartesianYOSGB36 As Double
        Dim CartesianZOSGB36 As Double
        Dim CartesianXWGS84 As Double
        Dim CartesianYWGS84 As Double
        Dim CartesianZWGS84 As Double

        Dim DX As Double = -446.448 'translation parallel to X 
        Dim DY As Double = 125.157 'translation parallel to Y 
        Dim DZ As Double = -542.06 'translation parallel to Z 
        Dim s As Double = 20.4894 'scale change 
        Dim X_Rot As Double = -0.1502 'rotation about X 
        Dim Y_Rot As Double = -0.247 'rotation about Y 
        Dim Z_Rot As Double = -0.8421 'rotation about Z 

        'Ellipsoid parameters WGS84
        Dim aWGS84 As Double = 6378137.0
        Dim bWGS84 As Double = 6356752.314

        CartesianXWGS84 = Lat_Long_H_to_X(LatWGS84, LonWGS84, ODNHeight, aWGS84, bWGS84)
        CartesianYWGS84 = Lat_Long_H_to_Y(LatWGS84, LonWGS84, ODNHeight, aWGS84, bWGS84)
        CartesianZWGS84 = Lat_H_to_Z(LatWGS84, ODNHeight, aWGS84, bWGS84)

        CartesianXOSGB36 = Helmert_X(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DX, Y_Rot, Z_Rot, s)
        CartesianYOSGB36 = Helmert_Y(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DY, X_Rot, Z_Rot, s)
        CartesianZOSGB36 = Helmert_Z(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DZ, X_Rot, Y_Rot, s)

        WGS84Long2OSGGB36 = XYZ_to_Long(CartesianXOSGB36, CartesianYOSGB36)

    End Function

    Public Function WGS84Lat2OSGGB36(ByVal LatWGS84 As Double, ByVal LongWGS84 As Double, ByVal ODNHeight As Double) As Double

        Dim CartesianXOSGB36 As Double
        Dim CartesianYOSGB36 As Double
        Dim CartesianZOSGB36 As Double
        Dim CartesianXWGS84 As Double
        Dim CartesianYWGS84 As Double
        Dim CartesianZWGS84 As Double

        Dim DX As Double = -446.448 'translation parallel to X 
        Dim DY As Double = 125.157 'translation parallel to Y 
        Dim DZ As Double = -542.06 'translation parallel to Z 
        Dim s As Double = 20.4894 'scale change 
        Dim X_Rot As Double = -0.1502 'rotation about X 
        Dim Y_Rot As Double = -0.247 'rotation about Y 
        Dim Z_Rot As Double = -0.8421 'rotation about Z 

        'Ellipsoid parameters WGS84
        Dim aWGS84 As Double = 6378137.0
        Dim bWGS84 As Double = 6356752.314

        CartesianXWGS84 = Lat_Long_H_to_X(LatWGS84, LongWGS84, ODNHeight, aWGS84, bWGS84)
        CartesianYWGS84 = Lat_Long_H_to_Y(LatWGS84, LongWGS84, ODNHeight, aWGS84, bWGS84)
        CartesianZWGS84 = Lat_H_to_Z(LatWGS84, ODNHeight, aWGS84, bWGS84)

        CartesianXOSGB36 = Helmert_X(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DX, Y_Rot, Z_Rot, s)
        CartesianYOSGB36 = Helmert_Y(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DY, X_Rot, Z_Rot, s)
        CartesianZOSGB36 = Helmert_Z(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DZ, X_Rot, Y_Rot, s)

        'Ellipsoid parameters WGS84
        Dim aOSGB36 As Double = 6377563.396
        Dim bOSGB36 As Double = 6356256.91

        WGS84Lat2OSGGB36 = XYZ_to_Lat(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, aOSGB36, bOSGB36)
    End Function

    Function Easting2Long(ByVal East As Double, ByVal North As Double) As Double
        'This is OSGB36
        Easting2Long = E_N_to_Long(East, North, a, b, e0, n0, f0, PHI0, LAM0)
    End Function

    Function Northing2Lat(ByVal East As Double, ByVal North As Double) As Double
        'This is OSGB36
        Northing2Lat = E_N_to_Lat(East, North, a, b, e0, n0, f0, PHI0, LAM0)
    End Function

    Public Function Lat2Northing(ByVal Lat As Double, ByVal Lon As Double) As Double
        'This is OSGB36
        Lat2Northing = Lat_Long_to_North(Lat, Lon, a, b, e0, n0, f0, PHI0, LAM0)
    End Function

    Public Function Lon2Easting(ByVal Lat As Double, ByVal Lon As Double) As Double
        'This is OSGB36
        Lon2Easting = Lat_Long_to_East(Lat, Lon, a, b, e0, f0, PHI0, LAM0)
    End Function

    Sub ParseInput(ByVal bAllLevels As Boolean)

        Dim i As Integer

        s100 = Left(sGridRef, 2)

        If sRefType = "100km" Then
            sHectad = ""
            sTetrad = ""
            sMonad = ""
            s6Fig = ""
            s8Fig = ""
            s10Fig = ""
            sEastingLL = Prefix2Easting(sGridRef) & "00000"
            sNorthingLL = Prefix2Northing(sGridRef) & "00000"
            sEastingC = Prefix2Easting(sGridRef) & "50000"
            sNorthingC = Prefix2Northing(sGridRef) & "50000"

            zoomLevel = 7
        End If

        If sRefType = "hectad" Then

            sHectad = sGridRef
            sTetrad = ""
            sMonad = ""
            s6Fig = ""
            s8Fig = ""
            s10Fig = ""
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 1) & "0000"
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 4, 1) & "0000"
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 1) & "5000"
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 4, 1) & "5000"

            zoomLevel = 11
        End If

        If sRefType = "quadrant" Then

            sHectad = Left(sGridRef, 4)
            sQuadrant(0) = sGridRef
            sTetrad = ""
            sMonad = ""
            s6Fig = ""
            s8Fig = ""
            s10Fig = ""
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 1) & QuadSuf2DX(Right(sGridRef, 2), False)
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 4, 1) & QuadSuf2DY(Right(sGridRef, 2), False)
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 1) & QuadSuf2DX(Right(sGridRef, 2), True)
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 4, 1) & QuadSuf2DY(Right(sGridRef, 2), True)

            zoomLevel = 12
        End If

        If sRefType = "tetrad" Then

            Dim sSplitString As String
            Dim sQuadrantSuffixes() As String

            sHectad = Left(sGridRef, 4)
            sSplitString = TetSuf2QuadSuf(Right(sGridRef, 1))
            sQuadrantSuffixes = sSplitString.Split("/")
            For i = 0 To UBound(sQuadrantSuffixes)
                sQuadrant(i) = Left(sGridRef, 4) & sQuadrantSuffixes(i)
            Next

            sTetrad = sGridRef
            sMonad = ""
            s6Fig = ""
            s8Fig = ""
            s10Fig = ""
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 1) & TetSuf2DX(Right(sGridRef, 1), False)
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 4, 1) & TetSuf2DY(Right(sGridRef, 1), False)
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 1) & TetSuf2DX(Right(sGridRef, 1), True)
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 4, 1) & TetSuf2DY(Right(sGridRef, 1), True)

            zoomLevel = 13
        End If

        If sRefType = "monad" Then

            sHectad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 5, 1)
            sQuadrant(0) = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 5, 1) & Mon2QuadSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 6, 1))
            sTetrad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 5, 1) & Mon2TetSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 6, 1))
            sMonad = sGridRef
            s6Fig = ""
            s8Fig = ""
            s10Fig = ""
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 2) & "000"
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 5, 2) & "000"
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 2) & "500"
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 5, 2) & "500"

            zoomLevel = 14
        End If

        If sRefType = "6fig" Then

            sHectad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 6, 1)
            sQuadrant(0) = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 6, 1) & Mon2QuadSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 7, 1))
            sTetrad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 6, 1) & Mon2TetSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 7, 1))
            sMonad = Left(sGridRef, 2) & Mid(sGridRef, 3, 2) & Mid(sGridRef, 6, 2)
            s6Fig = sGridRef
            s8Fig = ""
            s10Fig = ""
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 3) & "00"
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 6, 3) & "00"
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 3) & "50"
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 6, 3) & "50"

            zoomLevel = 15
        End If

        If sRefType = "8fig" Then

            sHectad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 7, 1)
            sQuadrant(0) = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 7, 1) & Mon2QuadSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 8, 1))
            sTetrad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 7, 1) & Mon2TetSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 8, 1))
            sMonad = Left(sGridRef, 2) & Mid(sGridRef, 3, 2) & Mid(sGridRef, 7, 2)
            s6Fig = Left(sGridRef, 2) & Mid(sGridRef, 3, 3) & Mid(sGridRef, 7, 3)
            s8Fig = sGridRef
            s10Fig = ""
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 4) & "0"
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 7, 4) & "0"
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 4) & "5"
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 7, 4) & "5"

            zoomLevel = 17
        End If

        If sRefType = "10fig" Then

            sHectad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 8, 1)
            sQuadrant(0) = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 8, 1) & Mon2QuadSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 9, 1))
            sTetrad = Left(sGridRef, 2) & Mid(sGridRef, 3, 1) & Mid(sGridRef, 8, 1) & Mon2TetSuf(Mid(sGridRef, 4, 1), Mid(sGridRef, 9, 1))
            sMonad = Left(sGridRef, 2) & Mid(sGridRef, 3, 2) & Mid(sGridRef, 8, 2)
            s6Fig = Left(sGridRef, 2) & Mid(sGridRef, 3, 3) & Mid(sGridRef, 8, 3)
            s8Fig = Left(sGridRef, 2) & Mid(sGridRef, 3, 4) & Mid(sGridRef, 8, 4)
            s10Fig = sGridRef
            sEastingLL = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 5)
            sNorthingLL = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 8, 5)
            sEastingC = Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 5) & ".5"
            sNorthingC = Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 8, 5) & ".5"

            zoomLevel = 19
        End If

        '100km Square

        If bAllLevels Or sRefType = "100km" Then
            East = CDbl(Prefix2Easting(s100) & "50000")
            North = CDbl(Prefix2Northing(s100) & "50000")
            fWidth = 50000
        End If

        If (sHectad <> "" And bAllLevels) Or sRefType = "hectad" Then

            East = CDbl(Prefix2Easting(Left(sHectad, 2)) & Mid(sHectad, 3, 1) & "5000")
            North = CDbl(Prefix2Northing(Left(sHectad, 2)) & Mid(sHectad, 4, 1) & "5000")
            fWidth = 5000
        End If
        If (sQuadrant(0) <> "" And bAllLevels) Or sRefType = "quadrant" Then

            For i = 0 To 4
                If sQuadrant(i) > "" Then
                    East = CDbl(Prefix2Easting(Left(sQuadrant(i), 2)) & Mid(sQuadrant(i), 3, 1) & QuadSuf2DX(Right(sQuadrant(i), 2), True))
                    North = CDbl(Prefix2Northing(Left(sQuadrant(i), 2)) & Mid(sQuadrant(i), 4, 1) & QuadSuf2DY(Right(sQuadrant(i), 2), True))
                    fWidth = 2500
                End If
            Next

        End If
        If (sTetrad <> "" And bAllLevels) Or sRefType = "tetrad" Then

            East = CDbl(Prefix2Easting(Left(sTetrad, 2)) & Mid(sTetrad, 3, 1) & TetSuf2DX(Right(sTetrad, 1), True))
            North = CDbl(Prefix2Northing(Left(sTetrad, 2)) & Mid(sTetrad, 4, 1) & TetSuf2DY(Right(sTetrad, 1), True))
            fWidth = 1000
        End If
        If (sMonad <> "" And bAllLevels) Or sRefType = "monad" Then

            East = CDbl(Prefix2Easting(Left(sMonad, 2)) & Mid(sMonad, 3, 2) & "500")
            North = CDbl(Prefix2Northing(Left(sMonad, 2)) & Mid(sMonad, 5, 2) & "500")
            fWidth = 500
        End If
        If (s6Fig <> "" And bAllLevels) Or sRefType = "6fig" Then

            East = CDbl(Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 3) & "50")
            North = CDbl(Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 6, 3) & "50")
            fWidth = 50
        End If
        If (s8Fig <> "" And bAllLevels) Or sRefType = "8fig" Then

            East = CDbl(Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 4) & "5")
            North = CDbl(Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 7, 4) & "5")
            fWidth = 5
        End If

        If (s10Fig <> "" And bAllLevels) Or sRefType = "10fig" Then

            East = CDbl(Prefix2Easting(Left(sGridRef, 2)) & Mid(sGridRef, 3, 5) & ".5")
            North = CDbl(Prefix2Northing(Left(sGridRef, 2)) & Mid(sGridRef, 8, 5) & ".5")
            fWidth = 0.5
        End If

        East = CDbl(sEastingC)
        North = CDbl(sNorthingC)
    End Sub

    Function LongWGS842Easting(ByVal LatWGS84 As Double, ByVal LongWGS84 As Double, ByVal EllipsoidalHeight As Double) As Double

        Dim CartesianXWGS84 As Double
        Dim CartesianYWGS84 As Double
        Dim CartesianZWGS84 As Double
        Dim CartesianXOSGB36 As Double
        Dim CartesianYOSGB36 As Double
        Dim CartesianZOSGB36 As Double
        Dim LongOSGB36 As Double
        Dim LatOSGB36 As Double

        'Ellipsoid parameters WGS84
        Dim aWGS84 As Double = 6378137.0
        Dim bWGS84 As Double = 6356752.314

        'Transformation (WGS84 to OSGB36)
        Dim DX As Double = -446.448 'translation parallel to X 
        Dim DY As Double = 125.157 'translation parallel to Y 
        Dim DZ As Double = -542.06 'translation parallel to Z 
        Dim s As Double = 20.4894 'scale change 
        Dim X_Rot As Double = -0.1502 'rotation about X 
        Dim Y_Rot As Double = -0.247 'rotation about Y 
        Dim Z_Rot As Double = -0.8421 'rotation about Z 

        CartesianXWGS84 = Lat_Long_H_to_X(LatWGS84, LongWGS84, EllipsoidalHeight, aWGS84, bWGS84)
        CartesianYWGS84 = Lat_Long_H_to_Y(LatWGS84, LongWGS84, EllipsoidalHeight, aWGS84, bWGS84)
        CartesianZWGS84 = Lat_H_to_Z(LatWGS84, EllipsoidalHeight, aWGS84, bWGS84)

        CartesianXOSGB36 = Helmert_X(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DX, Y_Rot, Z_Rot, s)
        CartesianYOSGB36 = Helmert_Y(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DY, X_Rot, Z_Rot, s)
        CartesianZOSGB36 = Helmert_Z(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DZ, X_Rot, Y_Rot, s)

        LatOSGB36 = XYZ_to_Lat(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, a, b)
        LongOSGB36 = XYZ_to_Long(CartesianXOSGB36, CartesianYOSGB36)

        LongWGS842Easting = Lon2Easting(LatOSGB36, LongOSGB36)
    End Function

    Function LatWGS842Northing(ByVal LatWGS84 As Double, ByVal LongWGS84 As Double, ByVal EllipsoidalHeight As Double) As Double

        Dim CartesianXWGS84 As Double
        Dim CartesianYWGS84 As Double
        Dim CartesianZWGS84 As Double
        Dim CartesianXOSGB36 As Double
        Dim CartesianYOSGB36 As Double
        Dim CartesianZOSGB36 As Double
        Dim LongOSGB36 As Double
        Dim LatOSGB36 As Double

        'Ellipsoid parameters WGS84
        Dim aWGS84 As Double = 6378137.0
        Dim bWGS84 As Double = 6356752.314

        'Transformation (WGS84 to OSGB36)
        Dim DX As Double = -446.448 'translation parallel to X 
        Dim DY As Double = 125.157 'translation parallel to Y 
        Dim DZ As Double = -542.06 'translation parallel to Z 
        Dim s As Double = 20.4894 'scale change 
        Dim X_Rot As Double = -0.1502 'rotation about X 
        Dim Y_Rot As Double = -0.247 'rotation about Y 
        Dim Z_Rot As Double = -0.8421 'rotation about Z 

        CartesianXWGS84 = Lat_Long_H_to_X(LatWGS84, LongWGS84, EllipsoidalHeight, aWGS84, bWGS84)
        CartesianYWGS84 = Lat_Long_H_to_Y(LatWGS84, LongWGS84, EllipsoidalHeight, aWGS84, bWGS84)
        CartesianZWGS84 = Lat_H_to_Z(LatWGS84, EllipsoidalHeight, aWGS84, bWGS84)

        CartesianXOSGB36 = Helmert_X(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DX, Y_Rot, Z_Rot, s)
        CartesianYOSGB36 = Helmert_Y(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DY, X_Rot, Z_Rot, s)
        CartesianZOSGB36 = Helmert_Z(CartesianXWGS84, CartesianYWGS84, CartesianZWGS84, DZ, X_Rot, Y_Rot, s)

        LatOSGB36 = XYZ_to_Lat(CartesianXOSGB36, CartesianYOSGB36, CartesianZOSGB36, a, b)
        LongOSGB36 = XYZ_to_Long(CartesianXOSGB36, CartesianYOSGB36)

        LatWGS842Northing = Lat2Northing(LatOSGB36, LongOSGB36)
    End Function
End Class
