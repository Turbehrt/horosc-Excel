Attribute VB_Name = "Sequences"

' This module is part of the application Horosc for Excel (https://github.com/Turbehrt/horosc-Excel),
' based on John D. North's HOROSC software, after the MS-DOS Pascal code published in
' John D. North, Horoscopes and History (London: The Warburg Institute, 1986), Appendix 4,
' as well as its adaptation for Google Sheets by François J. Tur and Alexandre Tur, 2021-.

' Horosc for Excel is an adaptation in VBA for Microsoft Excel by François J. Tur and Alexandre Tur, 2025-.

' This software is governed by the CeCILL-B license under French law and
' abiding by the rules of distribution of free software.  You can  use,
' modify and/ or redistribute the software under the terms of the CeCILL-B
' license as circulated by CEA, CNRS and INRIA at the following URL:
' http://www.cecill.info.

Public Const ERROR_INVALID_METHOD As Long = vbObjectError + 514
Public Const ERROR_INVALID_RANGE As Long = vbObjectError + 515
Public Const ERROR_INVALID_INDEX As Long = vbObjectError + 516


' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6), based on any method (0-6)
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} geographical latitude of the observation location in radian
' @param {real} right ascension of the Ascendant in radian
' @param {real} longitude of the Ascendant in radian
' @param {real} right ascension of the IMC in radian
' @param {real} longitude of the IMC in radian
' @param {integer} index of the house to compute (1-6)
' @param {integer} index of the method to use (0-6)
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
 
Function computeCuspWithMethodInRadian(ByVal obliquity As Double, ByVal geoLat As Double, ByVal rightASC As Double, ByVal longASC As Double, ByVal rightIMC As Double, ByVal longIMC As Double, ByVal houseIndex As Integer, ByVal method As Integer, Optional ByVal getRA As Boolean = True) As Double
    ' Compute a cusp given a known obliquity, geographical latitude, ascendant, IMC, and method
    
    'TODO: houses 7 to 12 could be computed by symetry (this was not in the original programme)
    If houseIndex > 6 Then
        Err.Raise ERROR_INVALID_METHOD, "computeCuspWithMethodInRadian", "HouseIndex must be an integer between 1 and 6"
    End If
    Dim cusp As Double
    
    Select Case method
        Case 0
            cusp = Method0(obliquity, geoLat, rightASC, rightIMC, houseIndex, getRA)
        Case 1
            cusp = Method1(obliquity, rightASC, rightIMC, houseIndex, getRA)
        Case 2
            cusp = Method2(obliquity, longASC, longIMC, houseIndex, getRA)
        Case 3
            cusp = Method3(obliquity, geoLat, rightIMC, longIMC, houseIndex, getRA)
        Case 4
            cusp = Method4(obliquity, geoLat, rightIMC, longIMC, houseIndex, getRA)
        Case 5
            cusp = Method5(obliquity, rightASC, houseIndex, getRA)
        Case 6
            cusp = Method6(obliquity, longASC, houseIndex, getRA)
        Case Else
            Err.Raise ERROR_INVALID_METHOD, "computeCuspWithMethodInRadian", "Method must be an integer between 0 and 6"
    End Select
    
   computeCuspWithMethodInRadian = cusp
        
End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6), for any method (0-6), based on geographical latitude and ascendant longitude
' (coordinates of the IMC are computed from the ascensional difference, cf. North's formula 5)
'
' @param {real} obliquity in radian
' @param {real} geographical latitude of the obervation location in radian
' @param {real} longitude of the Ascendant in radian
' @param {integer} index of the house to compute (1 though 6)
' @param {integer} index of the method to use (0 though 6)
' @param {boolean} which coordinate to return (true = right ascention, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates
' @customfunction
Function computeCuspFromLatitudeInRadian(ByVal obliquity As Double, ByVal geoLat As Double, ByVal longASC As Double, ByVal houseIndex As Integer, ByVal method As Integer, Optional ByVal getRA As Boolean = True) As Double
    ' Compute a cusp given a known obliquity, geographical latitude, ascendant, and method
    
    ' Pre-computes IMC (rightIMC = rightVernalPoint + pi/2)
    ' Formula (5) : sin(ascensionalDifference) = tan(obliquity) * tan(geoLat) * sin(rightASC)
    ' where the ascensionalDifference = rightASC - rightVernalPoint

    Dim rightASC As Double, rightIMC As Double, longIMC As Double, ascensionalDifference As Double
    
    rightASC = EclipticToEquator(obliquity, longASC)
    
    ascensionalDifference = WorksheetFunction.Asin(Sin(rightASC) * Tan(obliquity) * Tan(geoLat))
   
    rightIMC = rightASC + WorksheetFunction.Pi() / 2 - ascensionalDifference
    longIMC = EquatorToEcliptic(obliquity, rightIMC)
    
    computeCuspFromLatitudeInRadian = computeCuspWithMethodInRadian(obliquity, geoLat, rightASC, longASC, rightIMC, longIMC, houseIndex, method, getRA)
            
End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6), for any method (0-6), based on geographical latitude and ascendant longitude, all expressed in sexagesimal degrees.
'
' @param {string} obliquity in sexagesimal
' @param {string} geographical latitude of the observation location in sexagesimal
' @param {string} longitude of the Ascendant in sexagesimal
' @param {integer} index of the house to compute (1-6)
' @param {integer} index of the method to use (0-6)
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {string} the computed coordinate in sexagesimal - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function computeCuspFromLatitudeInSexagesimal(ByVal obliquitySxg As String, ByVal geoLatSxg As String, ByVal longASCSxg As String, ByVal houseIndex As Integer, ByVal method As Integer, Optional ByVal getRA As Boolean = True) As String

    Dim obliquity As Double, longASC As Double, geoLat As Double, cusp As Double
    
    obliquity = SexagesimalToRadian(obliquitySxg)
    longASC = SexagesimalToRadian(longASCSxg)
    geoLat = SexagesimalToRadian(geoLatSxg)
    
    cusp = computeCuspFromLatitudeInRadian(obliquity, geoLat, longASC, houseIndex, method, getRA)
    
    computeCuspFromLatitudeInSexagesimal = RadianToSexagesimal(cusp, SexagesimalFormat(longASCSxg))

End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6), for any method (0-6) based on observed longitudes.
' (geographical latitude is not given but deduced from the right ascensions of the Ascendant and IMC)
'
' @param {real} obliquity in radian
' @param {real} observed longitude of the Ascendant in radian
' @param {real} observed longitude of the IMC in radian
' @param {integer} index of the house to compute (1-6)
' @param {integer} index of the method to use (0-6)
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function computeCuspFromLongitudeInRadian(ByVal obliquity As Double, ByVal longASC As Double, ByVal longIMC As Double, ByVal houseIndex As Integer, ByVal method As Integer, Optional ByVal getRA As Boolean = True) As Double
    ' Compute a cusp given a known obliquity, ascendant, MC, and method (geographical latitude unknown)
    
    Dim rightASC As Double, rightIMC As Double, geoLat As Double
    
    rightASC = EclipticToEquator(obliquity, longASC)
    rightIMC = EclipticToEquator(obliquity, longIMC)
    geoLat = RetrieveLatitude(obliquity, rightASC, rightIMC)
    
    computeCuspFromLongitudeInRadian = computeCuspWithMethodInRadian(obliquity, geoLat, rightASC, longASC, rightIMC, longIMC, houseIndex, method, getRA)
            
End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6), for any method (0-6), based on observed longitudes, all expressed in sexagesimal degrees.
'
' @param {string} obliquity in sexagesimal
' @param {string} observed longitude of the Ascendant in sexagesimal
' @param {string} observed longitude of the IMC in sexagesimal
' @param {integer} index of the house to compute (1-6)
' @param {integer} index of the method to use (0-6)
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {string} the computed coordinate in sexagesimal - this function must be called twice to get both coordinates (right ascention and longitude)
' @customfunction
Function computeCuspFromLongitudeInSexagesimal(ByVal obliquitySxg As String, ByVal longASCSxg As String, ByVal longIMCSxg As String, ByVal houseIndex As Integer, ByVal method As Integer, Optional ByVal getRA As Boolean = True) As String

    Dim obliquity As Double, longASC As Double, longIMC As Double, cusp As Double
    
    obliquity = SexagesimalToRadian(obliquitySxg)
    longASC = SexagesimalToRadian(longASCSxg)
    longIMC = SexagesimalToRadian(longIMCSxg)
    
    cusp = computeCuspFromLongitudeInRadian(obliquity, longASC, longIMC, houseIndex, method, getRA)
    
    computeCuspFromLongitudeInSexagesimal = RadianToSexagesimal(cusp, SexagesimalFormat(longASCSxg))
             
End Function


' helper function for below all-in-one table computation
' copy data from one array to another
Function moveRows(ByRef dest() As Variant, ByRef source() As Variant, ByVal destFirstLine As Integer)
   Dim addedRows As Double
   addedRows = 0
    For r = LBound(source) To UBound(source)
        For c = LBound(source, 2) To UBound(source, 2)
            dest(r + destFirstLine, c) = source(r, c)
        Next c
        addedRows = addedRows + 1
    Next r
    moveRows = destFirstLine + addedRows
End Function

' helper function for below all-in-one table computation
' fill dest rows with empty string(instead of empty rows)
Function fillEmpty(ByRef dest() As Variant, ByVal CountLines As Integer, ByVal destFirstLine As Integer)
   Dim addedRows As Double
   addedRows = 0
    For r = 1 To CountLines
        For c = LBound(dest, 2) To UBound(dest, 2)
            dest(r + destFirstLine, c) = ""
        Next c
        addedRows = addedRows + 1
    Next r
    fillEmpty = destFirstLine + addedRows
End Function

' Method A : computes the theoretical calculation of the sexagesimal longitudes and/or right ascensions of the first 6 houses (the next 6 are inferred by symmetry) according to the 7 historical house systems
' @param {string} obliquity of the ecliptic in sexagesimal
' @param {string} geographical latitude of observation location in sexagesimal
' @param {string} longitude of the Ascendant in sexagesimal
' @param {integer} number of rows to separate the 2 result tables
' @return {string} an array of lines and columns that provides all longitudes and right ascensions computed
' @customfunction

Public Function ComputeLongitudesAllMethodsLatitude( _
    ByVal obliquitySx As String, _
    ByVal geoLatitudeSx As String, _
    ByVal longASCSx As String, _
    ByVal nbLinesGap As Long _
) As Variant

    ' Maps to Google Script's 'const methods = [0, 1, 2, 3, 4, 5, 6];'
    Dim methods(0 To 6) As Long
    Dim i As Long
    For i = 0 To 6: methods(i) = i: Next i

    ' Maps to Google Script's 'const houses = [1, 2, 3, 4, 5, 6];'
    Dim houses(1 To 6) As Long
    For i = 1 To 6: houses(i) = i: Next i

    ' In VBA, array indices start at 1 for a 1-based output array in Excel.
    ' We need 6 rows (for houses 1 to 6) and 7 columns (for methods 0 to 6).
    Dim zoneAscendantsSx(1 To 6, 1 To 7) As Variant ' Right Ascensions
    Dim zoneLongitudesSx(1 To 6, 1 To 7) As Variant ' Longitudes
    
    Dim house As Integer
    Dim method As Integer
    Dim hInd As Long ' Row index (1 to 6)
    Dim mInd As Long ' Column index (1 to 7)

    ' Equivalent to the nested forEach loops
    For hInd = 1 To 6
        house = houses(hInd)
        For mInd = 1 To 7
            method = methods(mInd - 1) ' methods array is 0-based, output array is 1-based

            ' Compute Right Ascension (ascendant) - The 6th parameter is TRUE by default in the original, but here we enforce it for clarity.
            zoneAscendantsSx(hInd, mInd) = computeCuspFromLatitudeInSexagesimal( _
                obliquitySx, geoLatitudeSx, longASCSx, house, method, True)

            ' Compute Longitude
            zoneLongitudesSx(hInd, mInd) = computeCuspFromLatitudeInSexagesimal( _
                obliquitySx, geoLatitudeSx, longASCSx, house, method, False)

        Next mInd
    Next hInd

    ' Maps to: [...zoneLongitudesSx, ...Array(nbLinesGap).fill(""), ...zoneAscendantsSx]
    ' Determine the size of the final output array
    Dim totalRows As Long, lastRow As Integer
    Dim resultArray() As Variant
    
    totalRows = 6 + nbLinesGap + 6
    ReDim resultArray(1 To totalRows, 1 To 7) As Variant
    
    lastRow = 0
    lastRow = moveRows(resultArray, zoneLongitudesSx, lastRow)
    lastRow = fillEmpty(resultArray, nbLinesGap, lastRow)
    lastRow = moveRows(resultArray, zoneAscendantsSx, lastRow)
    

    ' Return the final array to Excel
    ComputeLongitudesAllMethodsLatitude = resultArray

End Function

' Method B: Theoretical calculation of the latitude of the observation site (with an interval corresponding to the margin of error, applied to the right ascension of the ascendant or the midheaven) and a comparison with the theoretical longitudes (calculated considering only the ascendant and the midheaven) according to the seven historical house systems, with a quality coefficient (generally allowing the method actually used to be identified)
' @param {string} obliquity of the ecliptic in sexagesimal
' @param {range of strings} observed longitudes for all 6 houses in sexagesimal
' @param {string} error for geographical latitude deviation in sexagesimal
' @param {integer} number of rows to separate each result table
' @return {string} an array of lines and columns that provides all theoretical longitudes, quality coefficients, theoretical right ascensions, and the cross of geographical latitude deviation
' @customfunction

Function computeLongitudesAllMethodsLongitude( _
  ByVal obliquitySx As String, _
  ByVal longitudesSx As Variant, _
  ByVal errorLgSx As String, _
  ByVal nbLinesGap As Integer) As Variant

  ' Maps to Google Script's 'const methods = [0, 1, 2, 3, 4, 5, 6];'
  Dim methods(0 To 6) As Long
  Dim i As Long
  For i = 0 To 6: methods(i) = i: Next i

  ' Maps to Google Script's 'const houses = [1, 2, 3, 4, 5, 6];'
  Dim houses(1 To 6) As Long
  For i = 1 To 6: houses(i) = i: Next i

  ' Maps to Google Script's 'const qualities = [1, 2, 3, 4, 5, 6];'
  Dim qualities(1 To 6) As Long
  For i = 1 To 6: qualities(i) = i: Next i

    
  'const zoneAscendantsSx = houses.map((L) => Array(methods.length).fill("-"));
  'const zoneLongitudesSx = houses.map((L) => Array(methods.length).fill("-"));
  'const zoneQualityDeg = qualities.map((L) => Array(methods.length).fill("-"));
  'const zoneAvgQualityDeg = [Array(methods.length).fill("-")];

    Dim zoneAscendantsSx(1 To 6, 1 To 7) As Variant ' Right Ascensions
    Dim zoneLongitudesSx(1 To 6, 1 To 7) As Variant ' Longitudes
    Dim zoneQualityDeg(1 To 6, 1 To 7) As Variant 'Qualities
    Dim zoneAvgQualityDeg(1 To 1, 1 To 7) As Variant 'Average quality for each method
    Dim displayLatitudeApproxSx(1 To 3, 1 To 3) As Variant
    
    
    Dim house As Integer
    Dim method As Integer
    Dim hInd As Long ' Row index (1 to 6)
    Dim mInd As Long ' Column index (1 to 7)

    Dim longASCSx As String, longIMCSx As String
    
   longASCSx = longitudesSx(1) 'by definition the ASC is the observed longiude of house 1
   longIMCSx = longitudesSx(4) 'by definition the IMC is the observed longiude of house 4


    ' Equivalent to the nested forEach loops
    For hInd = 1 To 6
        house = houses(hInd)
        For mInd = 1 To 7
            method = methods(mInd - 1) ' methods array is 0-based, output array is 1-based

            ' **NOTE:** Assumes computeCuspFromLatitudeInSexagesimal is a translated VBA function
            ' Compute Right Ascension (ascendant) - The 6th parameter is TRUE by default in the original, but here we enforce it for clarity.
            zoneAscendantsSx(hInd, mInd) = computeCuspFromLongitudeInSexagesimal( _
                obliquitySx, longASCSx, longIMCSx, house, method, True)

            ' Compute Longitude
            zoneLongitudesSx(hInd, mInd) = computeCuspFromLongitudeInSexagesimal( _
                obliquitySx, longASCSx, longIMCSx, house, method, False)
            
            zoneQualityDeg(hInd, mInd) = QualityCoefficientDegree(longitudesSx(hInd), zoneLongitudesSx(hInd, mInd))

        Next mInd
    Next hInd

    For mInd = 1 To 7
        zoneAvgQualityDeg(1, mInd) = (zoneQualityDeg(2, mInd) + zoneQualityDeg(3, mInd) + zoneQualityDeg(5, mInd) + zoneQualityDeg(6, mInd)) / 4
    Next mInd
    displayLatitudeApproxSx(1, 1) = ""
    displayLatitudeApproxSx(1, 2) = RetrieveLatitudeRangeSexagesimal(obliquitySx, longASCSx, longIMCSx, errorLgSx, 3)
    displayLatitudeApproxSx(1, 3) = ""
    displayLatitudeApproxSx(2, 1) = RetrieveLatitudeRangeSexagesimal(obliquitySx, longASCSx, longIMCSx, errorLgSx, 1)
    displayLatitudeApproxSx(2, 2) = RetrieveLatitudeRangeSexagesimal(obliquitySx, longASCSx, longIMCSx, errorLgSx, 0)
    displayLatitudeApproxSx(2, 3) = RetrieveLatitudeRangeSexagesimal(obliquitySx, longASCSx, longIMCSx, errorLgSx, 2)
    displayLatitudeApproxSx(3, 1) = ""
    displayLatitudeApproxSx(3, 2) = RetrieveLatitudeRangeSexagesimal(obliquitySx, longASCSx, longIMCSx, errorLgSx, 4)
    displayLatitudeApproxSx(3, 3) = ""
    
        ' Maps to: [...zoneLongitudesSx, ...Array(nbLinesGap).fill(""), ...zoneAscendantsSx]
    ' Determine the size of the final output array
    Dim totalRows As Long, lastRow As Integer
    Dim resultArray() As Variant
    
    totalRows = 6 + nbLinesGap + 1 + 6 + nbLinesGap + 6 + nbLinesGap + 3
    ReDim resultArray(1 To totalRows, 1 To 7) As Variant

    lastRow = 0
    lastRow = moveRows(resultArray, zoneLongitudesSx, lastRow)
    lastRow = fillEmpty(resultArray, nbLinesGap, lastRow)
    lastRow = moveRows(resultArray, zoneAvgQualityDeg, lastRow)
    lastRow = moveRows(resultArray, zoneQualityDeg, lastRow)
    lastRow = fillEmpty(resultArray, nbLinesGap, lastRow)
    lastRow = moveRows(resultArray, zoneAscendantsSx, lastRow)
    lastRow = fillEmpty(resultArray, nbLinesGap + 3, lastRow)
    lastRow = moveRows(resultArray, displayLatitudeApproxSx, lastRow - 3)

    ' Return the final array to Excel
    computeLongitudesAllMethodsLongitude = resultArray

End Function

