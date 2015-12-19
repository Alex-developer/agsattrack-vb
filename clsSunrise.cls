VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSunrise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ------ Class clsSunRiseSet

Option Explicit

' -- The following properties are exposed:
'Sunrise (r) - Sunrise time
'Sunset (r) - Sunset time
'SolarNoon (r) - Solar noon
'
'CityCount (r) - Number of cities
'CityName (r) - Name of city, by index
'City (w) - Sets the longitude/latitude/timezone based off a city
'           name or city index
'
'TimeZone (r/w) - Current Timezone
'DaySavings (r/w) - Daylight savings time in effect
'Longitude (r/w) - Longitude to calculate for
'Latitude (r/w) - Latitude to calculate for
'
'DateDay (r/w) - Date to calculate for
'
'
' -- The following method is exposed
'CalculateSun - Calculate sunrise, sunset and solar noon
'

Private Type typeMonth
    Name As String
    NumDays As Long
End Type

Private Type typeCity
    Name As String
    Longitude As Double
    Latitude As Double
    TimeZone As Long
End Type

Private m_cNumberCities As Long
Private m_Cities() As typeCity
    
Private m_monthList(0 To 11) As typeMonth
Private m_monthLeap(0 To 11) As typeMonth

Private m_nTimeZone As Long
Private m_bDaySavings As Boolean
Private m_nLongitude As Double
Private m_nLatitude As Double
Private m_dateSel As Date

Private m_dateSunrise As Date
Private m_dateSunset As Date
Private m_dateNoon As Date

Public Property Get Sunrise() As Date
    Sunrise = m_dateSunrise
End Property

Public Property Get Sunset() As Date
    Sunset = m_dateSunset
End Property

Public Property Get SolarNoon() As Date
    SolarNoon = m_dateNoon
End Property

Public Property Get CityCount() As Long
    CityCount = m_cNumberCities + 1
End Property

Public Property Get CityName(nCity As Long) As String
    If nCity < 0 Or nCity > m_cNumberCities Then
        CityName = "(Error)"
    Else
        CityName = m_Cities(nCity).Name
    End If
End Property

Public Property Let City(City)
    Dim nCity As Long
    Dim bFound As Boolean
    
    If VarType(City) = vbString Then
        For nCity = 0 To m_cNumberCities
            If Trim(LCase(City)) = Trim(LCase(m_Cities(nCity).Name)) Then
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            nCity = -1
        End If
    Else
        If IsNumeric(City) Then
            nCity = City
        Else
            nCity = -1
        End If
    End If
    
    If nCity < 0 Or nCity > m_cNumberCities Then
        m_nTimeZone = 0
        m_bDaySavings = False
        m_nLongitude = 0
        m_nLatitude = 0
    Else
        m_nTimeZone = m_Cities(nCity).TimeZone
        m_bDaySavings = False
        m_nLongitude = m_Cities(nCity).Longitude
        m_nLatitude = m_Cities(nCity).Latitude
    End If
    
End Property

Public Property Let TimeZone(nNew As Long)
    m_nTimeZone = nNew
End Property

Public Property Get TimeZone() As Long
    TimeZone = m_nTimeZone
End Property

Public Property Let DaySavings(bNew As Boolean)
    m_bDaySavings = bNew
End Property

Public Property Get DaySavings() As Boolean
    DaySavings = m_bDaySavings
End Property

Public Property Let Longitude(nNew As Double)
    m_nLongitude = nNew
End Property

Public Property Get Longitude() As Double
    Longitude = m_nLongitude
End Property

Public Property Let Latitude(nNew As Double)
    m_nLatitude = nNew
End Property

Public Property Get Latitude() As Double
    Latitude = m_nLatitude
End Property
    
Public Property Let DateDay(dateNew As Date)
    m_dateSel = dateNew
End Property
    
Public Property Get DateDay() As Date
    DateDay = m_dateSel
End Property


Private Function IsLeapYear(nYear As Long) As Boolean
    If (nYear Mod 4 = 0 And nYear Mod 100 <> 0) Or nYear Mod 400 = 0 Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If
End Function

Private Function RadToDeg(angleRad As Double) As Double
    RadToDeg = 180 * angleRad / 3.1415926535
End Function

Private Function DegToRad(angleDeg As Double) As Double
    DegToRad = 3.1415926535 * angleDeg / 180
End Function

Private Function CalcJulianDay(nMonth As Long, nDay As Long, bLeapYear _
   As Boolean) As Long


    Dim i As Long
    Dim nJulianDay As Long
    
    
    If bLeapYear Then
        For i = 0 To nMonth - 1
            nJulianDay = nJulianDay + m_monthLeap(i).NumDays
        Next
    Else
        For i = 0 To nMonth - 1
            nJulianDay = nJulianDay + m_monthList(i).NumDays
        Next
    End If
    
    nJulianDay = nJulianDay + nDay
    
    CalcJulianDay = nJulianDay
End Function

Private Function CalcGamma(nJulianDay As Long) As Double

    CalcGamma = (2 * 3.1415926535 / 365) * (nJulianDay - 1)

End Function

Private Function CalcGamma2(nJulianDay As Long, hour As Long)

    CalcGamma2 = (2 * 3.1415926535 / 365) * (nJulianDay - 1 + (hour / 24))

End Function

Private Function CalcEqOfTime(gamma As Double) As Double

    CalcEqOfTime = (229.18 * (0.000075 + 0.001568 * Cos(gamma) - _
       0.032077 * Sin(gamma) - 0.014615 * Cos(2 * gamma) - 0.040849 * _
       Sin(2 * gamma)))

    
End Function

Private Function CalcSolarDec(gamma As Double) As Double

    CalcSolarDec = (0.006918 - 0.399912 * Cos(gamma) + 0.070257 * _
       Sin(gamma) - 0.006758 * Cos(2 * gamma) + 0.000907 * Sin(2 * _
       gamma))


End Function

Private Function acos(x As Double) As Double
    On Error Resume Next
    acos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Private Function CalcHourAngle(lat As Double, solarDec As Double, time _
   As Boolean) As Double

    Dim latRad As Double
    latRad = DegToRad(lat)

    If time Then
        CalcHourAngle = (acos(Cos(DegToRad(90.833)) / (Cos(latRad) * _
           Cos(solarDec)) - Tan(latRad) * Tan(solarDec)))

    Else
        CalcHourAngle = -(acos(Cos(DegToRad(90.833)) / (Cos(latRad) * _
           Cos(solarDec)) - Tan(latRad) * Tan(solarDec)))

    End If

End Function

Private Function CalcDayLength(hourAngle As Double) As Double
    CalcDayLength = (2 * Abs(RadToDeg(hourAngle))) / 15
End Function

Public Function CalculateSun()

    Dim nLatitude As Double
    Dim nLongitude As Double
    Dim dateCalc As Date
    Dim bDaySavings As Long
    Dim nZone As Long
    
    nLatitude = m_nLatitude
    nLongitude = m_nLongitude
    dateCalc = m_dateSel
    bDaySavings = IIf(m_bDaySavings, 60, 0)
    nZone = m_nTimeZone

    If nLatitude >= -90 And nLatitude < -89.8 Then
        nLatitude = -89.8
    End If
    If nLatitude <= 90 And nLatitude > 89.8 Then
        nLatitude = 89.8
    End If
    
    ' Calculate the time of sunrise
    
    Dim nJulianDay As Long
    nJulianDay = CalcJulianDay(Month(dateCalc), Day(dateCalc), _
       IsLeapYear(Year(dateCalc)))

    
    Dim gamma_solnoon As Double
    Dim eqTime As Double
    Dim solarDec As Double
    
    gamma_solnoon = CalcGamma2(nJulianDay, 12 + (nLongitude / 15))
    eqTime = CalcEqOfTime(gamma_solnoon)
    solarDec = CalcSolarDec(gamma_solnoon)
    
    Dim timeGMT As Double
    timeGMT = CalcSunriseGMT(DatePart("y", dateCalc), nLatitude, _
       nLongitude)

    
    Dim solNoonGmt As Double
    solNoonGmt = CalcSolNoonGMT(DatePart("y", dateCalc), nLongitude)
    
    Dim timeLST As Double
    timeLST = timeGMT - (60 * nZone) + bDaySavings
    m_dateSunrise = DateAdd("s", timeLST * 60, Int(m_dateSel))
    
    '*****Calculate Solar noon
    Dim solNoonLST As Double
    solNoonLST = solNoonGmt - (60 * nZone) + bDaySavings
    m_dateNoon = DateAdd("s", solNoonLST * 60, Int(m_dateSel))
    
    
    '***** Calculate the time of sunset
    Dim setTimeGMT As Double
    Dim setTimeLST As Double
    
    setTimeGMT = CalcSunsetGMT(DatePart("y", dateCalc), nLatitude, _
       nLongitude)

    setTimeLST = setTimeGMT - (60 * nZone) + bDaySavings
    m_dateSunset = DateAdd("s", setTimeLST * 60, Int(m_dateSel))
    
End Function

Private Function CalcSunriseGMT(nJulianDay As Long, nLatitude As Double, _
   nLongitude As Double)


    Dim gamma As Double
    Dim eqTime As Double
    Dim solarDec As Double
    Dim hourAngle As Double
    Dim delta As Double
    Dim timeDiff As Double
    Dim timeGMT As Double
    
    gamma = CalcGamma(nJulianDay)
    eqTime = CalcEqOfTime(gamma)
    solarDec = CalcSolarDec(gamma)
    hourAngle = CalcHourAngle(nLatitude, solarDec, True)
    delta = nLongitude - RadToDeg(hourAngle)
    timeDiff = 4 * delta
    timeGMT = 720 + timeDiff - eqTime
    
    Dim gamma_sunrise As Double
    
    gamma_sunrise = CalcGamma2(nJulianDay, timeGMT / 60)
    eqTime = CalcEqOfTime(gamma_sunrise)
    solarDec = CalcSolarDec(gamma_sunrise)
    hourAngle = CalcHourAngle(nLatitude, solarDec, True)
    delta = nLongitude - RadToDeg(hourAngle)
    timeDiff = 4 * delta
    timeGMT = 720 + timeDiff - eqTime
    
    CalcSunriseGMT = timeGMT

End Function

Private Function CalcSolNoonGMT(nJulianDay As Long, nLongitude As _
   Double) As Double


    Dim gamma_solnoon As Double
    Dim eqTime As Double
    Dim solarNoonDec As Double
    Dim solNoonGmt As Double
    
    gamma_solnoon = CalcGamma2(nJulianDay, 12 + (nLongitude / 15))
    eqTime = CalcEqOfTime(gamma_solnoon)
    solarNoonDec = CalcSolarDec(gamma_solnoon)
    solNoonGmt = 720 + (nLongitude * 4) - eqTime
    
    CalcSolNoonGMT = solNoonGmt

End Function

Private Function CalcSunsetGMT(nJulianDay As Long, nLatitude As Double, _
   nLongitude As Double) As Double


    Dim gamma As Double
    Dim eqTime As Double
    Dim solarDec As Double
    Dim hourAngle As Double
    Dim delta As Double
    Dim timeDiff As Double
    Dim setTimeGMT As Double
    
    gamma = CalcGamma(nJulianDay + 1)
    eqTime = CalcEqOfTime(gamma)
    solarDec = CalcSolarDec(gamma)
    hourAngle = CalcHourAngle(nLatitude, solarDec, False)
    delta = nLongitude - RadToDeg(hourAngle)
    timeDiff = 4 * delta
    setTimeGMT = 720 + timeDiff - eqTime
    
    Dim gamma_sunset As Double
    
    gamma_sunset = CalcGamma2(nJulianDay, setTimeGMT / 60)
    eqTime = CalcEqOfTime(gamma_sunset)
    
    solarDec = CalcSolarDec(gamma_sunset)
    
    hourAngle = CalcHourAngle(nLatitude, solarDec, False)
    delta = nLongitude - RadToDeg(hourAngle)
    timeDiff = 4 * delta
    setTimeGMT = 720 + timeDiff - eqTime
    
    CalcSunsetGMT = setTimeGMT

End Function

Private Sub InitMonths()

    m_monthList(0).Name = "January": m_monthList(0).NumDays = 31
    m_monthList(1).Name = "February": m_monthList(1).NumDays = 28
    m_monthList(2).Name = "March": m_monthList(2).NumDays = 31
    m_monthList(3).Name = "April": m_monthList(3).NumDays = 30
    m_monthList(4).Name = "May": m_monthList(4).NumDays = 31
    m_monthList(5).Name = "June": m_monthList(5).NumDays = 30
    m_monthList(6).Name = "July": m_monthList(6).NumDays = 31
    m_monthList(7).Name = "August": m_monthList(7).NumDays = 31
    m_monthList(8).Name = "September": m_monthList(8).NumDays = 30
    m_monthList(9).Name = "October": m_monthList(9).NumDays = 31
    m_monthList(10).Name = "November": m_monthList(10).NumDays = 30
    m_monthList(11).Name = "DEcember": m_monthList(11).NumDays = 31
    
    m_monthLeap(0).Name = "January": m_monthLeap(0).NumDays = 31
    m_monthLeap(1).Name = "February": m_monthLeap(1).NumDays = 28
    m_monthLeap(2).Name = "March": m_monthLeap(2).NumDays = 31
    m_monthLeap(3).Name = "April": m_monthLeap(3).NumDays = 30
    m_monthLeap(4).Name = "May": m_monthLeap(4).NumDays = 31
    m_monthLeap(5).Name = "June": m_monthLeap(5).NumDays = 30
    m_monthLeap(6).Name = "July": m_monthLeap(6).NumDays = 31
    m_monthLeap(7).Name = "August": m_monthLeap(7).NumDays = 31
    m_monthLeap(8).Name = "September": m_monthLeap(8).NumDays = 30
    m_monthLeap(9).Name = "October": m_monthLeap(9).NumDays = 31
    m_monthLeap(10).Name = "November": m_monthLeap(10).NumDays = 30
    m_monthLeap(11).Name = "DEcember": m_monthLeap(11).NumDays = 31

End Sub

Private Sub AddCity(sCity As String, nLatitude As Double, nLongitude As _
   Double, nZone As Long)


    m_cNumberCities = m_cNumberCities + 1
    If m_cNumberCities > UBound(m_Cities) Then
        ReDim Preserve m_Cities(UBound(m_Cities) + 10)
    End If

    m_Cities(m_cNumberCities).Name = sCity
    m_Cities(m_cNumberCities).Latitude = nLatitude
    m_Cities(m_cNumberCities).Longitude = nLongitude
    m_Cities(m_cNumberCities).TimeZone = nZone

End Sub

Private Sub InitCities()

    m_cNumberCities = -1
    ReDim Preserve m_Cities(0)

    AddCity "Albuquerque, NM", 35.05, 106.39, 7
    AddCity "Anchorage, AK", 61.13, 149.54, 9
    AddCity "Atlanta, GA", 33.44, 84.23, 5
    AddCity "Boston, MA", 42.21, 71.03, 5
    AddCity "Boulder, CO", 40.125, 105.237, 7
    AddCity "Chicago, IL", 41.51, 87.39, 6
    AddCity "Dallas, TX", 32.46, 96.47, 6
    AddCity "Denver, CO", 39.44, 104.59, 7
    AddCity "Detroit, MI", 42.2, 83.03, 5
    AddCity "Honolulu, HA", 21.18, 157.51, 10
    AddCity "Indianapolis, IN", 39.46, 86.09, 5
    AddCity "Kansas City, MO", 39.05, 94.34, 6
    AddCity "Los Angeles, CA", 34.03, 118.14, 8
    AddCity "Miami, FL", 25.46, 80.11, 5
    AddCity "Minneapolis, nMonth", 44.58, 93.15, 6
    AddCity "New Orleans, LA", 29.57, 90.04, 6
    AddCity "New York City, NY", 40.43, 74.01, 5
    AddCity "Oklahoma City, OK", 35.28, 97.3, 6
    AddCity "Philadelphia, PA", 39.57, 75.09, 5
    AddCity "Phoenix, AZ", 33.26, 112.04, 7
    AddCity "Saint Louis, MO", 38.37, 90.11, 6
    AddCity "San Fransisco, CA", 37.46, 122.25, 8
    AddCity "Seattle, WA", 47.36, 122.19, 8
    AddCity "Washington DC", 38.53, 77.02, 5
    AddCity "Beijing, China", 39.55, -116.25, -8
    AddCity "Berlin, Germany", 52.33, -13.3, -1
    AddCity "Buenos Aires, Argentina", -34.36, 58.27, 3
    AddCity "Cairo, Egypt", -30.06, -31.22, -2
    AddCity "Cape Town, South Africa", -33.55, -18.22, -2
    AddCity "Caracas, Venezula", 10.3, 66.56, 4
    AddCity "Helsinki, Finland", 60.1, -24.58, -2
    AddCity "Hong Kong, China", 22.15, -114.1, -8
    AddCity "London, England", 51.3, 0.1, 0
    AddCity "Mexico City, Mexico", 19.24, 99.09, 6
    AddCity "Moscow, Russia", 55.45, -37.35, -3
    AddCity "New Delhi, India", 28.36, -77.12, -5.5
    AddCity "Ottawa, Canada", 45.25, 75.42, 5
    AddCity "Paris, France", 48.52, -2.2, -1
    AddCity "Rio de Janeiro, Brazil", -22.54, 43.14, 3
    AddCity "Riyadh, Saudi Arabia", 24.38, -46.43, -3
    AddCity "Rome, Italy", 41.54, -12.29, -1
    AddCity "Sydney, Australia", -33.52, -151.13, -10
    AddCity "Tokyo, Japan", 35.42, -139.46, -9
    AddCity "Zurich, Switzerland", 47.23, -8.32, -1

End Sub

Private Sub Class_Initialize()

    InitMonths
    InitCities
    m_dateSel = Now

End Sub

' ------ End of class clsSunRiseSet

