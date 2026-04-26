Attribute VB_Name = "Domification"

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

' Convention names :
' obliquity is the obliquity of the ecliptic
' geoLat is the geographical latitude of the place (usually phi), equivalent to the elevation of the poles (over the horizon)
' ascensionalDifference is the ascensional difference (right ascendant - right ascension of the vernal point)
' Any domification method computes the cusps of each House (1 to 6, with houses 7 to 12 being symetrical).
    ' they come as right Ascensions (on the equinox, usually noted alpha) or longitudes (on the ecliptic, usually noted lambda).
' Meaningful cusps are the Ascendant (Asc, cusp of House 1, intersection of ecliptic and horizon) and the Immum Caeli (IMC, cusp of House 4, intersection of ecliptic and night meridian)
    ' right ascension of the ascendant - ascensional difference = right ascension of the IMC - 90°
' Houses 7 to 12 are symetrical to Houses  to 6 in all methods (cusp + 180)

Const middleSkyCuspIndex = 4

' converts a longitude in right ascension for a given obliquity of the ecliptic
' @param {real} obliquity in radian
' @param {real} longitude in radian
' @return {real} the computed right ascension
' @customfunction

Function EclipticToEquator(ByVal obliquity As Double, ByVal longitude As Double) As Double    'ECEQ in Pascal
    ' Formula 2 : tan(rightAscension) = cos(obliquity) * tan (longitude)
    
    Dim rightAscension As Double
    ' rightAscension = Atn(Cos(obliquity) * Sin(longitude) / Cos(longitude))
    rightAscension = WorksheetFunction.Atan2(Cos(longitude), Cos(obliquity) * Sin(longitude))
    EclipticToEquator = ModuloTwoPI(rightAscension)
    
End Function

' converts a right ascension in longitude for a given obliquity of the ecliptic
' @param {real} obliquity in radian
' @param {real} right ascension in radian
' @return {real} the computed longitude
' @customfunction

Function EquatorToEcliptic(ByVal obliquity As Double, ByVal rightAscension As Double) As Double    'EQEC in Pascal
    ' Formula 2 : tan(rightAscension) = cos(obliquity) * tan (longitude) <=> tan (longitude) = tan(rightAscension) / cos(obliquity)
    
    Dim longitude As Double
    ' longitude = Atn(Sin(rightAscension) / (Cos(obliquity) * Cos(rightAscension)))
    longitude = WorksheetFunction.Atan2(Cos(rightAscension), Sin(rightAscension) / Cos(obliquity))
    
    'TODO: instruction below is useless(tested with all RA values)
    ' If Abs(longitude - rightAscension) > WorksheetFunction.Pi() / 5# Then longitude = longitude + WorksheetFunction.Pi()
    
    EquatorToEcliptic = ModuloTwoPI(longitude)
    
End Function

' calculates the theoretical latitude of the observation location from the obliquity of the ecliptic and the right ascension of the ascendant and the Imum Caeli (IMC, opposite the Medium Caeli).
' @param {real} obliquity in radian
' @param {real} right ascension of the ascendant in radian
' @param {real} right ascension of the IMC in radian
' @return {real} the geographical latitude in radian
' @customfunction
Function RetrieveLatitude(ByVal obliquity As Double, ByVal rightASC As Double, ByVal rightIMC As Double, Optional ByVal methNorth As Boolean = True)
  
  Dim bracket As Double, lat As Double
  bracket = rightIMC - rightASC
  lat = WorksheetFunction.Atan2(Sin(rightASC) * Tan(obliquity), Cos(bracket))
  
  If methNorth Then
  ' Changed the modulo above from 2Pi to 1Pi to match the symetry to Pi/2
    lat = ModuloTwoPI(Abs(lat))
    If (lat > WorksheetFunction.Pi() / 2) And (lat < WorksheetFunction.Pi()) Then
        lat = WorksheetFunction.Pi() - lat
    End If
    
  ' Experimental : allows for negative latitudes (for Southern hemisphere)
  Else
    If lat > WorksheetFunction.Pi() / 2 Then lat = lat - WorksheetFunction.Pi()
    If lat < -WorksheetFunction.Pi() / 2 Then lat = lat + WorksheetFunction.Pi()
  End If
  RetrieveLatitude = lat
  
End Function

' calculates the theoretical latitude of the observation location from the obliquity of the ecliptic and the longitude of the ascendant and the Imum Caeli (IMC).
' @param {real} obliquity in radian
' @param {real} longitude of the ascendant in radian
' @param {real} longitude of the IMC in radian
' @return {real} the computed geographical latitude in radian
' @customfunction
Function RetrieveLatitudeFromLong(ByVal obliquity As Double, ByVal longASC As Double, ByVal longIMC As Double, Optional ByVal methNorth As Boolean = True)
  Dim rigthASC As Double, rightIMC As Double
  
  rightASC = EclipticToEquator(obliquity, longASC)
  rightIMC = EclipticToEquator(obliquity, longIMC)
  
  RetrieveLatitudeFromLong = RetrieveLatitude(obliquity, rightASC, rightIMC, methNorth)
  
End Function


' calculates the theoretical latitude of the observation location from the right ascensions of the ascendant and the Imum Caeli (IMC) and the obliquity of the ecliptic, all expressed in sexagesimal degrees.
'
' @param {string} obliquity of the ecliptic in sexagesimal
' @param {string} right ascension of the Ascendant in sexagesimal
' @param {string} right ascension of the IMC in sexagesimal
' @return {string} the computed geographical latitude of observation location in sexagesimal
' @customfunction
Function RetrieveLatitudeSexagesimal(ByVal obliquitySxg As String, ByVal rightASCSxg As String, ByVal rightIMCSxg As String, Optional ByVal methNorth As Boolean = True) As String
    Dim obliquity As Double, rightASC As Double, rightIMC As Double
    
    ' Sexagesimal to Radian
    obliquity = SexagesimalToRadian(obliquitySxg)
    rightASC = SexagesimalToRadian(rightASCSxg)
    rightIMC = SexagesimalToRadian(rightIMCSxg)
   
    RetrieveLatitudeSexagesimal = RadianToSexagesimal(RetrieveLatitude(obliquity, rightASC, rightIMC, methNorth), SexagesimalFormat(rightASCSxg))
End Function

' calculates the theoretical latitude of the observation location from the longitudes of the ascendant and the Imum Caeli (IMC) and the obliquity of the ecliptic, all expressed in sexagesimal degrees.
'
' @param {string} obliquity of the ecliptic in sexagesimal
' @param {string} longitude of the Ascendant in sexagesimal
' @param {string} longitude of the IMC in sexagesimal
' @return {string} the computed geographical latitude of observation location in sexagesimal
' @customfunction
Function RetrieveLatitudeFromLongSexagesimal(ByVal obliquitySxg As String, ByVal longASCSxg As String, ByVal longIMCSxg As String, Optional ByVal methNorth As Boolean = True) As String
    Dim obliquity As Double, longASC As Double, longIMC As Double
    Dim rightASC As Double, rightIMC As Double
    
    ' Sexagesimal to Radian
    obliquity = SexagesimalToRadian(obliquitySxg)
    longASC = SexagesimalToRadian(longASCSxg)
    longIMC = SexagesimalToRadian(longIMCSxg)
   
    RetrieveLatitudeFromLongSexagesimal = RadianToSexagesimal(RetrieveLatitudeFromLong(obliquity, longASC, longIMC, methNorth), SexagesimalFormat(longASCSxg))
End Function

' calculates the theoretical latitude of the observation location from the obliquity of the ecliptic and the longitudes of the ascendant and the Imum Caeli (IMC) in radian, when deviated of an error of observation (or rounding approximation) in a direction
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} longitude of the Ascendant in radian
' @param {real} longitude of the IMC in radian
' @param {real} error of observation in radian
' @param {integer} direction of the error - a value from 0 to 4.
'   - 0 indicates no deviation
'   - 1 and 2 indicate deviation by error value respectively in lower or excess from the longitude of the ascendant
'   - 4 and 3 indicate deviation by error value respectively in lower or excess from the longitude of the IMC/MC
' @return {real} the computed geographical latitude in radian
' @customfunction

Function RetrieveLatitudeRange(ByVal obliquity As Double, ByVal longASC As Double, ByVal longIMC As Double, ByVal error As Double, ByVal direction As Integer) As Double
    Dim rightASC As Double, rightIMC As Double, distance As Double
    Dim latitude As Double

    ' Right ascensions
    rightASC = EclipticToEquator(obliquity, longASC)
    rightIMC = EclipticToEquator(obliquity, longIMC)
    
    Select Case direction
    Case 0
        ' 0 Exact value
        RetrieveLatitudeRange = RetrieveLatitude(obliquity, rightASC, rightIMC)
        
    ' Case 1 (uppper) and 2 (lower) : assumption that the astrologer started from the ascendent (+/- max error)
    Case 1
        ' 1 rightASC - error
        ' NOTE: it is a change from initial program
        ' Initial program: FOI = RetrieveLatitudeRange = retrieveLatitude(obliquity, rightASC - error, rightIMC - error)
        RetrieveLatitudeRange = RetrieveLatitude(obliquity, rightASC - error, rightIMC)
    Case 2
        ' 2 rightASC + error
        ' NOTE: it is a change from initial program
        ' Initial program: FIO = RetrieveLatitudeRange = retrieveLatitude(obliquity, rightASC + error, rightIMC - error)
        RetrieveLatitudeRange = RetrieveLatitude(obliquity, rightASC + error, rightIMC)
    
    ' Case 3 (left) and 4 (right) : assumption that the astrologer started from the MC (+/- max error)
    Case 3
        ' 3 rightIMC + error
        RetrieveLatitudeRange = RetrieveLatitude(obliquity, rightASC, rightIMC + error)
    Case 4
        ' 4 rightIMC - error
        RetrieveLatitudeRange = RetrieveLatitude(obliquity, rightASC, rightIMC - error)
        
    Case Else
        Err.Raise ERROR_INVALID_METHOD, "RetrieveLatitudeRange", "range value must be between 0 and 4"
    End Select
    
End Function

' calculates the theoretical latitude of the observation location from the obliquity of the ecliptic and the longitudes of the ascendant and the Imum Caeli (IMC) in sexagesimal degrees, when deviated of an error of observation (or rounding approximation) in a direction
'
' @param {string} obliquity of the ecliptic in sexagesimal
' @param {string} longitude of the Ascendant in sexagesimal
' @param {string} longitude of the IMC in sexagesimal
' @param {string} error of observation in sexagesimal
' @param {integer} direction of the error - a valude from 0 to 4.
'   - 0 indicated no deviation
'   - 1 and 2 indicate deviation by error value respectively in lower or excess from longitude of ascendant
'   - 4 and 3 indicate deviation by error value respectively in lower or excess from longitude of IMC
' @return {string} the computed geographical latitude in sexagesimal
' @customfunction
Function RetrieveLatitudeRangeSexagesimal(ByVal obliquitySxg As String, ByVal longASCSxg As String, ByVal longIMCSxg As String, ByVal errorSxg As String, ByVal direction As Integer) As String

    RetrieveLatitudeRangeSexagesimal = _
        RadianToSexagesimal( _
            RetrieveLatitudeRange(SexagesimalToRadian(obliquitySxg), _
                SexagesimalToRadian(longASCSxg), _
                SexagesimalToRadian(longIMCSxg), _
                SexagesimalToRadian(errorSxg), _
                direction _
            ), _
            SexagesimalFormat(longASCSxg) _
        )
        
End Function

Private Function Converge(ByVal cusp As Double, ByVal midCusp As Double, ByVal cuspNumber As Integer, ByVal obliquity As Double, ByVal geoLat As Double, Optional ByVal maxError As Double = 0.000002) As Double
    Dim errorVal As Double: errorVal = maxError + 1
    Dim v As Double: v = cusp
    Dim bb As Double, bracket As Double
    
    Do While errorVal >= maxError
        bracket = WorksheetFunction.Acos(Sin(v) * Tan(obliquity) * Tan(geoLat)) / 3#
        bb = ModuloTwoPI(midCusp - (4 - cuspNumber) * bracket)
        errorVal = Abs(bb - v)
        v = bb
    Loop
    
    Converge = v
End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 0:
' * Hour Lines (fixed boundaries) method
' * Cusps are intersections of the ecliptic by the horizon, the meridian circle, and the unequal hour lines on the sphere for even-numbered hours.
' * This method is usually graphical, with aid of an astrolabe (here emulated with a convergence function).
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} geographical latitude of the observation location in radian
' @param {real} right ascension of the Ascendant in radian
' @param {real} right ascension of the IMC in radian
' @param {integer} index of the house to compute (value 1 to 6)
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function Method0(ByVal obliquity As Double, ByVal geoLat As Double, ByVal rightASC As Double, ByVal rightIMC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double
    'Hour Lines (fixed boundaries) method
    'Cusps are intersections of the ecliptic by the horizon, the meridian circle, and the unequal hour lines on the sphere for even-numbered hours.
    'This method is usually graphical, with aid of an astrolabe. It is emulated using a convergence function.

    
    ' Initial bracket
    Dim b As Double, rightAscension As Double, longitude As Double
    
    b = ModuloTwoPI(rightIMC - rightASC) / 3#
    
    Select Case houseIndex
        Case 1
            rightAscension = rightASC
        Case 2
            rightAscension = Converge(ModuloTwoPI(rightASC + b), rightIMC, 2, obliquity, geoLat)
        Case 3
            rightAscension = Converge(ModuloTwoPI(rightASC + 2 * b), rightIMC, 3, obliquity, geoLat)
        Case 4
            rightAscension = rightIMC
        Case 5
            rightAscension = Converge(ModuloTwoPI(rightIMC + WorksheetFunction.Pi() / 3 - b), rightIMC, 5, obliquity, geoLat)
        Case 6
            rightAscension = Converge(ModuloTwoPI(rightIMC + 2 * (WorksheetFunction.Pi() / 3 - b)), rightIMC, 6, obliquity, geoLat)
        Case Else
            rightAscension = Null
    End Select
    longitude = EquatorToEcliptic(obliquity, rightAscension)
    If getRA Then Method0 = rightAscension Else Method0 = longitude
    
End Function


' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 1:
' * Standard method (Alcabitius)
' * Uniform division of of the cardinal sectors of the Equator
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} right ascension of the Ascendant in radian
' @param {real} right ascension of the IMC
' @param {integer} index of the house to compute
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function Method1(ByVal obliquity As Double, ByVal rightASC As Double, ByVal rightIMC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double
   
    Dim delta As Double, symDelta As Double, rightAscension As Double, longitude As Double
    
    delta = ModuloTwoPI(rightIMC - rightASC)
    symDelta = WorksheetFunction.Pi() - delta
    
       Select Case houseIndex
        Case 1
            rightAscension = rightASC
        Case 2
            rightAscension = ModuloTwoPI(rightASC + delta / 3)
        Case 3
            rightAscension = ModuloTwoPI(rightASC + 2 * delta / 3)
        Case 4
            rightAscension = rightIMC
        Case 5
            rightAscension = ModuloTwoPI(rightIMC + symDelta / 3)
        Case 6
            rightAscension = ModuloTwoPI(rightIMC + 2 * symDelta / 3)
        Case Else
            rightAscension = Null
    End Select
    longitude = EquatorToEcliptic(obliquity, rightAscension)
    
    If getRA Then Method1 = rightAscension Else Method1 = longitude

End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 2:
' * Dual longitude method
' * Uniform division of of the cardinal sectors of the Equator
'
' @param {real} obliquity od the ecliptic in radian
' @param {real} longitude of the Ascendant in radian
' @param {real} longitude of the IMC in radian
' @param {integer} index of the house to compute
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function Method2(ByVal obliquity As Double, ByVal longASC As Double, ByVal longIMC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double
    
    Dim delta As Double, symDelta As Double, longitude As Double, rightAscension As Double
    
    delta = ModuloTwoPI(longIMC - longASC)
    symDelta = WorksheetFunction.Pi() - delta
    
      Select Case houseIndex
        Case 1
            longitude = longASC
        Case 2
            longitude = ModuloTwoPI(longASC + delta / 3)
        Case 3
            longitude = ModuloTwoPI(longASC + 2 * delta / 3)
        Case 4
            longitude = longIMC
        Case 5
            longitude = ModuloTwoPI(longIMC + symDelta / 3)
        Case 6
            longitude = ModuloTwoPI(longIMC + 2 * symDelta / 3)
        Case Else
            longitude = Null
    End Select
    rightAscension = EclipticToEquator(obliquity, longitude)
    If getRA Then Method2 = rightAscension Else Method2 = longitude
       
End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 3:
' * Prime Vertical (fixed boundaries) method
' * Uniform division of the Prime Vertical
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} geographical latitude of the observation location in radian
' @param {real} right ascendant of the IMC in radian
' @param {real} longitude of the IMC in radian
' @param {integer} index of the house to compute
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude).
' @customfunction
Function Method3(ByVal obliquity As Double, ByVal geoLat As Double, ByVal rightIMC As Double, ByVal longIMC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double
    
  Dim rightAscension As Double, longitude As Double
    
    Select Case houseIndex
        Case 4
            rightAscension = rightIMC
            longitude = longIMC
        Case 1 To 3, 5 To 6
            Dim theta As Double, H As Double
             theta = (houseIndex - 1) * WorksheetFunction.Pi() / 6
            H = WorksheetFunction.Atan2(Cos(geoLat), Tan(theta))
            If H < 0 Then H = H + WorksheetFunction.Pi()
            
            rightAscension = ModuloTwoPI(rightIMC - WorksheetFunction.Pi() / 2 + H)
            longitude = ModuloTwoPI(WorksheetFunction.Atan2(Cos(rightAscension) * Cos(obliquity) - Cos(H) * Tan(geoLat) * Sin(obliquity), Sin(rightAscension)))
        Case Else
            rightAscension = Null
            longitude = Null
    End Select
    
    If getRA Then Method3 = rightAscension Else Method3 = longitude

End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 4:
' * Equatorial (fixed boundaries) method
' * Uniform division of the Equator (local sphere)
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} geographical latitude of the observation location in radian
' @param {real} right ascension of the IMC in radian
' @param {real} longitude of the IMC in radian
' @param {integer} index of the house to compute
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function Method4(ByVal obliquity As Double, ByVal geoLat As Double, ByVal rightIMC As Double, ByVal longIMC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double

    Dim rightAscension As Double, longitude As Double
    
    
    Select Case houseIndex
        Case 4
            rightAscension = rightIMC
            longitude = longIMC
        Case 1 To 3, 5 To 6
            Dim H As Double
            H = (houseIndex - 1) * WorksheetFunction.Pi() / 6
        
            rightAscension = ModuloTwoPI(rightIMC - WorksheetFunction.Pi() / 2 + H)
            longitude = ModuloTwoPI(WorksheetFunction.Atan2(Cos(rightAscension) * Cos(obliquity) - Cos(H) * Tan(geoLat) * Sin(obliquity), Sin(rightAscension)))
        Case Else
            rightAscension = Null
            longitude = Null
    End Select
    
    If getRA Then Method4 = rightAscension Else Method4 = longitude

End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 5:
' * Equatorial (moving boundaries) method
' * Uniform division of the Equator (celestial sphere)
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} right asension of the Ascendant in radian
' @param {integer} index of the house to compute
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function Method5(ByVal obliquity As Double, ByVal rightASC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double

   Dim rightAscension As Double, longitude As Double
    
    Select Case houseIndex
        Case 1 To 6
            rightAscension = ModuloTwoPI(rightASC + (houseIndex - 1) * WorksheetFunction.Pi() / 6)
        Case Else
            rightAscension = Null
    End Select
    longitude = EquatorToEcliptic(obliquity, rightAscension)
    
    If getRA Then Method5 = rightAscension Else Method5 = longitude
    
End Function

' computes the coordinates (right ascension, longitude) of the cusp of any house (1-6) following North's method 6:
' * Single longitudes method
' * Uniform division of the Ecliptic
'
' @param {real} obliquity of the ecliptic in radian
' @param {real} longitude of the Ascendant in radian
' @param {boolean} which coordinate to return (true = right ascension, false = longitude) - default true
' @return {real} the computed coordinate in radian - this function must be called twice to get both coordinates (right ascension and longitude)
' @customfunction
Function Method6(ByVal obliquity As Double, ByVal longASC As Double, ByVal houseIndex As Integer, Optional ByVal getRA As Boolean = True) As Double

  Dim longitude As Double, rightAscension As Double
    
    Select Case houseIndex
        Case 1 To 6
            longitude = ModuloTwoPI(longASC + (houseIndex - 1) * WorksheetFunction.Pi() / 6)
        Case Else
            longitude = Null
    End Select
    rightAscension = EclipticToEquator(obliquity, longitude)

    If getRA Then Method6 = rightAscension Else Method6 = longitude

End Function


' computes quality coefficient between observed longitude and computed longitude
'
' @param {real} observed longitude in radian
' @param {real} computed longitude in radian
' @return {real} quality coefficient in radian
' @customfunction
Function QualityCoefficientRadian(ByVal observedLongitude As Double, ByVal computedLongitude As Double) As Double
    ' Difference between the cusp provided and the expected value
    QualityCoefficientRadian = ModuloRange(Abs(computedLongitude - observedLongitude), WorksheetFunction.Pi())
End Function

' computes quality coefficient between observed longitude and computed longitude, expressed in sexagesimal degrees
' Note for comparison that in the original program, quality coefficient are expressed in radians
'
' @param {string} observed longitude in sexagesimal degrees
' @param {string} computed longitude in sexagesimal degrees
' @return {real} quality coefficient in decimal degrees
' @customfunction

Function QualityCoefficientDegree(ByVal observedLongitudeSxg As String, ByVal computedLongitudeSxg As String) As Double
    Dim observedLongitude As Double, computedLongitude As Double, coeff As Double
    
    observedLongitude = SexagesimalToRadian(observedLongitudeSxg)
    computedLongitude = SexagesimalToRadian(computedLongitudeSxg)
    coeff = QualityCoefficientRadian(observedLongitude, computedLongitude)
    QualityCoefficientDegree = coeff * (180# / WorksheetFunction.Pi())
End Function


