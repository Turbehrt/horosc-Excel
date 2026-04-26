Attribute VB_Name = "Trigonometry"

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

Public Const ERROR_INVALID_FMTCHAR As Long = vbObjectError + 530

' for a periodic function, ensures the value is within 0 - range, by cutting it by range size
' @param {real} value to bound
' @param {real} range width, by default the bound i [0, range].
' @param {boolean} optional indication if the bound should be centered on zero [-range, +range]
' @return {real} bounded value within range
' @customfunction

Function ModuloRange(ByVal value As Double, ByVal range As Double, Optional ByVal zCentered As Boolean = False) As Double

    Dim v As Double, s As Integer
    v = value
    s = 1
    If zCentered Then
        v = Abs(value)
        s = Sgn(value)
    Else
        Do While v < 0
            v = v + range
        Loop
    End If
    Do While v >= range
        v = v - range
    Loop
    ModuloRange = s * v
End Function

 ' returns the value within a range of 0-2PI (or -PI,+PI if centered)
 ' @customfunction

Function ModuloTwoPI(ByVal value As Double, Optional ByVal zCentered As Boolean = False) As Double
    ModuloTwoPI = ModuloRange(value, 2 * WorksheetFunction.Pi(), zCentered)
End Function

' Extracts the two triplets of values and delimiters from the sexagesimal value provided in parameters.
' Delimiters can be any char that are not numerical
' Left most value (the degrees) can be signed
' @param {sexa} a string showing a sexagesimal value. Eg - 90°30'00" or 90deg30m60s, etc.

Function SplitSexagesimalFormat(ByVal sexa As String) As Variant
    
    Dim delims(0 To 2) As String
    Dim values(0 To 2) As Double
    
    
    ' reads each char of the sexa value and extracts strings between numbers
    Dim i As Integer, part As Integer
    Dim currentValue As Double
    Dim currentDelim As String
    Dim currentSign As Boolean
    Dim valueInProgress As Boolean
    
    
    part = -1
    currentSign = False
    valueInProgress = False
    
    For i = 1 To Len(sexa)
        CHAR = Mid(sexa, i, 1)
            
        ' Check if character is NOT a digit (0-9) and NOT a minus sign (-)
        ' ASCII 48-57 are 0-9; ASCII 45 is minus
        If Not (CHAR >= "0" And CHAR <= "9") And Not (CHAR = "-" And Not valueInProgress) Then
            ' if value is in progress, record it
            If valueInProgress Then
              part = part + 1
              If currentSign Then currentValue = -currentValue
              values(part) = currentValue
              currentSign = False
              currentValue = 0#
              valueInProgress = False
            End If
            ' It's part of a delimiter, so builds it
            currentDelim = currentDelim & CHAR
        Else
            ' It IS a number/minus. If we were building a delimiter, saves it.
            ' If the delim is before the first number, we ignore it.
            If currentDelim <> "" And part >= 0 Then
                delims(part) = currentDelim
                currentDelim = ""
            End If
            valueInProgress = True
            If CHAR = "-" Then currentSign = True Else currentValue = currentValue * 10 + Val(CHAR)
        End If
    Next i
    
    ' After we process all, ensures to finish ongoing decoding
    If currentDelim <> "" And part < 3 Then
        delims(part) = currentDelim
    End If
    If valueInProgress And part < 2 Then
         If currentSign Then
             currentValue = -currentValue
        End If
        part = part + 1
        values(part) = currentValue
        delims(part) = ""
    End If
    
    For i = part + 1 To 2
       values(i) = 0#
       delims(i) = ""
    Next i
    
    Dim valuesAndDelims(0 To 1) As Variant
    valuesAndDelims(0) = values
    valuesAndDelims(1) = delims
    
    SplitSexagesimalFormat = valuesAndDelims

End Function

' Extracts the delimiters from the sexagesimal value.
' These delimiters can be passed as "frm" parameter for the function RadianToSexagesimal to format a radians value into sexagesimal.

Function SexagesimalFormat(ByVal sexa As String) As Variant

SexagesimalFormat = SplitSexagesimalFormat(sexa)(1)

End Function

' converts a sexagesimal string (representing an angle in degrees) into an angle in radian
' @param {string} sexa the sexagesimal value to convert
' @return {real} the equivalent radian value
' @customfunction

Function SexagesimalToRadian(ByVal sexa As String) As Double
    ' input: a string in three elements separated by one of: degree sign (°), dot, single-quote, space, double-quote, "d", semi-colon or any char provided in the seps parameter
    ' first part is the degrees, second part is minutes (0 to 59), third part is seconds (0 to 59)
    ' eg:134.30.25 - would mean 134 degrees and 30 minutes and 25 seconds
    ' return the corresponding value in radian.
    
    Dim parts() As Double
    parts = SplitSexagesimalFormat(sexa)(0)
    Dim r As Double, s As Double, v As Double
    r = 0#
    s = 1#
    For i = LBound(parts) To UBound(parts)
        v = parts(i)
        r = r + (v * s)
        s = s / 60#
    Next i
    SexagesimalToRadian = ModuloTwoPI((r / 360#) * 2 * WorksheetFunction.Pi(), True)
End Function

' converts an angle in radian into a sexagesimal representation in degrees, using a given set of symbols
' @param {real} radian value to convert
' @param {string or Array of strings} fmt - optional -  the format for presentation
'  - if provided as string: 2 chars, first one for degrees separator, second for minutes separator
'  - if provided as Array, then each element of the array are used as separator
' @return {string} the sexagesimal representation of the radian value
' @customfunction

Function RadianToSexagesimal(ByVal radian As Double, Optional ByVal frm As Variant) As String
   ' converts an arc valued in radian in the corresponding sexagesimal string representation: ddd°mm'ss"
   
    Dim dg As Double, d As Integer, m As Integer, s As Integer, sepDeg As String, sepMin As String, sepSeg As String, delims() As String
    sepDeg = ChrW(&HB0) 'degree symbol = U+00B0
    sepMin = "'"
    sepSec = """"
    
    dg = ModuloTwoPI(radian, True) * 360# / (2 * WorksheetFunction.Pi())
    d = Int(dg)
    'we need CInt to catch the round when diving doubles
    s = CInt((dg - d) * 3600)
    m = Int(s / 60)
    s = s - (m * 60)
    'because of CInt above, the m value can reach 60
    Do While m >= 60
       d = d + 1
       m = m - 60
    Loop
    
    If Not (IsMissing(frm)) Then
        If IsObject(frm) Then
            If VarType(frm) = vbString Then
                delims = SexagesimalFormat(frm)
            End If
        Else
            If IsArray(frm) Then
                delims = frm
            End If
        End If
        If UBound(delims) - LBound(delims) + 1 > 0 Then sepDeg = delims(LBound(delims)) Else sepDeg = "d"
        If UBound(delims) - LBound(delims) + 1 > 1 Then sepMin = delims(LBound(delims) + 1) Else sepMin = "'"
        If UBound(delims) - LBound(delims) + 1 > 2 Then sepSec = delims(LBound(delims) + 2) Else sepMin = """"
    End If
    
    RadianToSexagesimal = d & sepDeg & Format(m, "00") & sepMin & Format(s, "00") & sepSec
End Function
