VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSatellite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class allows a satellite to be tracked."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' **********************************************************************
' * Programmer Name  : Alex Greenland
' * Web Site         :
' * E-Mail           : greenlanda@aol.com
' * Date             : 03/11/99
' * Time             : 21:06
' * Module Name      : cSatellite
' * Module Filename  : cSatellite.cls
' **********************************************************************
' * Comments         : A Class module to track satellites
' *
' *
' **********************************************************************
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Const gc000C = 396 ' &H18C%
Private Const gc000E = 216 ' &HD8%
Private Const PixelsPerDegLon = 2.2
Private Const PixelsPerDegLat = 2.5
Private Const gc0020 = 800 ' &H320%

Private PI As Variant
Private MeanYear As Variant
Private TropicalYear As Variant
Private EarthRotationRate As Variant
Private EarthRotationRateDay As Variant
Private EarthRotationRateSeconds As Variant
Private SatDragCoeff As Variant
Private SatKepsMeanAnomoly As Variant
Private SatKepsMeanMotion As Variant
Private SatKepsRAAN As Variant
Private SatKepsArgOfPerigee As Variant

Private GravitationalConstant As Variant
Private ZonalCoeff As Variant
Private YG As Integer
Private G0 As Variant
Private LA As Variant
Private LO As Variant
Private HT As Variant
Private OLDRN As Integer
Private SatEpochDayNumber As Long
Private SatelliteTimeRequired As Double

Private RS(500) As Variant
Private CL As Variant
Private SL As Variant
Private CO As Variant
Private SO As Variant
Private RE As Variant
Private FL As Variant
Private RP As Variant
Private Rx As Variant
Private Ry As Variant
Private Rz As Variant
Private Ex As Variant
Private Ey As Variant
Private Ez As Variant
Private Sx As Variant
Private Sy As Variant
Private Sz As Variant
Private Ux As Variant
Private Uy As Variant
Private Uz As Variant
Private Ox As Variant
Private Oy As Variant
Private Oz As Variant
Private Ax As Variant
Private Ay As Variant
Private Az As Variant
Private Nx As Variant
Private Ny As Variant
Private Nz As Variant
Private VOx As Variant
Private VOy As Variant
Private M2 As Double
Private CI As Variant
Private SI As Variant
Private B0 As Variant
Private QD As Variant
Private WD As Variant
Private TEG As Variant
Private GHAE As Variant
Private N0 As Variant

'local variable(s) to hold property value(s)
Private mvarSatelliteName(500) As Variant 'local copy
Private mvarSatelliteDesignator(500) As Variant 'local copy
Private mvarSatelliteRange(500) As Long 'local copy
Private mvarSatelliteAzimuth(500) As Long 'local copy
Private mvarSatelliteElevation(500) As Long 'local copy
Private mvarKepsEccentricity(500) As Variant 'local copy
Private mvarKepsRAAN(500) As Variant 'local copy
Private mvarKepsAOP(500) As Variant 'local copy
Private mvarKepsInclination(500) As Variant 'local copy
Private mvarKepsMeanAnomoly(500) As Variant 'local copy
Private mvarKepsMeanMotion(500) As Variant 'local copy
Private mvarKepsOrbitNumber(500) As Variant 'local copy
Private mvarObserverLatitude As Variant 'local copy
Private mvarObserverLongitude As Variant 'local copy
Private mvarObserverHeight As Variant 'local copy
Private mvarKepsEpochTime(500) As Variant 'local copy
Private mvarKepsYearEpochTime(500) As Variant
Private mvarSatEpochYear(500) As Variant
Private mvarKepsElementSet(500) As Variant 'local copy
Private mvarKepsDecayRate(500) As Single 'local copy
Private mvarKepsChecksum(500) As Variant 'local copy
Private mvarSatelliteLatitude(500) As Long 'local copy
Private mvarSatelliteLongitude(500) As Long 'local copy
Private mvarSatelliteOrbitNumber(500) As Long 'local copy
Private mvarKepsYearEpochTimeFraction(500) As Variant
Private mvarSatelliteIndex As Integer 'local copy
Private mvarDisplayCentury(500) As Integer 'local copy
Private mvarDisplayYear(500) As Integer 'local copy
Private mvarDisplayMonth(500) As Integer 'local copy
Private mvarDisplayDay(500) As Integer 'local copy
Private mvarDisplayHour(500) As Integer 'local copy
Private mvarDisplayMinute(500) As Integer 'local copy
Private mvarSatelliteDayNumber(500) As Variant 'local copy
Private mvarDisplayTimeRequired(500) As Variant 'local copy
Private mvarSatelliteCount As Variant 'local copy
Private mvarDataValid(500) As Boolean 'local copy

Public Property Get DataValid() As Boolean
Attribute DataValid.VB_Description = "TRUE if data had been calculated for the selected satellite, FALSE if not."
    DataValid = mvarDataValid(mvarSatelliteIndex)
End Property

Public Property Get SatelliteCount() As Variant
  Dim i As Integer

  For i = 1 To 500
    If mvarSatelliteName(i) = "" Then Exit For
  Next i
  SatelliteCount = i - 1
End Property

Public Property Let DisplayTimeRequired(ByVal vData As Variant)
  mvarDisplayTimeRequired(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayTimeRequired() As Variant
  DisplayTimeRequired = mvarDisplayTimeRequired(mvarSatelliteIndex)
End Property

Public Property Let SatelliteDayNumber(ByVal vData As Variant)
  mvarSatelliteDayNumber(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteDayNumber() As Variant
  SatelliteDayNumber = mvarSatelliteDayNumber(mvarSatelliteIndex)
End Property

Public Property Let DisplayMinute(ByVal vData As Integer)
  mvarDisplayMinute(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayMinute() As Integer
  DisplayMinute = mvarDisplayMinute(mvarSatelliteIndex)
End Property

Public Property Let DisplayHour(ByVal vData As Integer)
  mvarDisplayHour(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayHour() As Integer
  DisplayHour = mvarDisplayHour(mvarSatelliteIndex)
End Property

Public Property Let DisplayDay(ByVal vData As Integer)
  mvarDisplayDay(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayDay() As Integer
  DisplayDay = mvarDisplayDay(mvarSatelliteIndex)
End Property

Public Property Let DisplayMonth(ByVal vData As Integer)
  mvarDisplayMonth(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayMonth() As Integer
  DisplayMonth = mvarDisplayMonth(mvarSatelliteIndex)
End Property

Public Property Let DisplayYear(ByVal vData As Integer)
  mvarDisplayYear(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayYear() As Integer
  DisplayYear = mvarDisplayYear(mvarSatelliteIndex)
End Property

Public Property Let DisplayCentury(ByVal vData As Integer)
  mvarDisplayCentury(mvarSatelliteIndex) = vData
End Property

Public Property Get DisplayCentury() As Integer
  DisplayCentury = mvarDisplayCentury(mvarSatelliteIndex)
End Property

Public Property Let SatelliteIndex(ByVal vData As Integer)
  mvarSatelliteIndex = vData
End Property

Public Property Get SatelliteIndex() As Integer
  SatelliteIndex = mvarSatelliteIndex
End Property

Public Property Let SatelliteOrbitNumber(ByVal vData As Long)
  mvarSatelliteOrbitNumber(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteOrbitNumber() As Long
  SatelliteOrbitNumber = mvarSatelliteOrbitNumber(mvarSatelliteIndex)
End Property

Public Property Let SatelliteLongitude(ByVal vData As Long)
  mvarSatelliteLongitude(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteLongitude() As Long
  SatelliteLongitude = mvarSatelliteLongitude(mvarSatelliteIndex)
End Property

Public Property Let SatelliteLatitude(ByVal vData As Long)
  mvarSatelliteLatitude(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteLatitude() As Long
  SatelliteLatitude = mvarSatelliteLatitude(mvarSatelliteIndex)
End Property

Public Property Let KepsChecksum(ByVal vData As Variant)
  mvarKepsChecksum(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsChecksum() As Variant
  KepsChecksum = mvarKepsChecksum(mvarSatelliteIndex)
End Property

Public Property Let KepsDecayRate(ByVal vData As Single)
  mvarKepsDecayRate(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsDecayRate() As Single
  KepsDecayRate = mvarKepsDecayRate(mvarSatelliteIndex)
End Property

Public Property Let KepsElementSet(ByVal vData As Variant)
  mvarKepsElementSet(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsElementSet() As Variant
  KepsElementSet = mvarKepsElementSet(mvarSatelliteIndex)
End Property

Public Property Let KepsEpochTime(ByVal vData As Variant)
  If Val(Left$(vData, 2)) < 50 Then
    century% = 20
  Else
    century% = 19
  End If
  mvarKepsEpochTime(mvarSatelliteIndex) = vData
  mvarKepsYearEpochTime(mvarSatelliteIndex) = mvarKepsEpochTime(mvarSatelliteIndex) - 1000 * Int(mvarKepsEpochTime(mvarSatelliteIndex) / 1000)
  mvarSatEpochYear(mvarSatelliteIndex) = 100 * century% + Int(mvarKepsEpochTime(mvarSatelliteIndex) / 1000)

End Property

Public Property Get KepsEpochTime() As Variant
  KepsEpochTime = mvarKepsEpochTime(mvarSatelliteIndex)
End Property

Public Property Let ObserverHeight(ByVal vData As Variant)
  mvarObserverHeight = vData / 1000
End Property

Public Property Get ObserverHeight() As Variant
  ObserverHeight = mvarObserverHeight * 1000
End Property

Public Property Let ObserverLongitude(ByVal vData As Variant)
  mvarObserverLongitude = FNRAD(vData)
End Property

Public Property Get ObserverLongitude() As Variant
  ObserverLongitude = FNIntDEG(mvarObserverLongitude)
End Property

Public Property Let ObserverLatitude(ByVal vData As Variant)
  mvarObserverLatitude = FNRAD(vData)
End Property

Public Property Get ObserverLatitude() As Variant
  ObserverLatitude = FNIntDEG(mvarObserverLatitude)
End Property

Public Property Let KepsOrbitNumber(ByVal vData As Variant)
  mvarKepsOrbitNumber(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsOrbitNumber() As Variant
  KepsOrbitNumber = mvarKepsOrbitNumber(mvarSatelliteIndex)
End Property

Public Property Let KepsMeanMotion(ByVal vData As Double)
  mvarKepsMeanMotion(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsMeanMotion() As Double
  KepsMeanMotion = mvarKepsMeanMotion(mvarSatelliteIndex)
End Property

Public Property Let KepsMeanAnomoly(ByVal vData As Single)
  mvarKepsMeanAnomoly(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsMeanAnomoly() As Single
  KepsMeanAnomoly = mvarKepsMeanAnomoly(mvarSatelliteIndex)
End Property

Public Property Let KepsInclination(ByVal vData As Single)
  mvarKepsInclination(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsInclination() As Single
  KepsInclination = mvarKepsInclination(mvarSatelliteIndex)
End Property

Public Property Let KepsAOP(ByVal vData As Single)
  mvarKepsAOP(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsAOP() As Single
  KepsAOP = mvarKepsAOP(mvarSatelliteIndex)
End Property

Public Property Let KepsRAAN(ByVal vData As Single)
  mvarKepsRAAN(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsRAAN() As Single
  KepsRAAN = mvarKepsRAAN(mvarSatelliteIndex)
End Property

Public Property Let KepsEccentricity(ByVal vData As Single)
  mvarKepsEccentricity(mvarSatelliteIndex) = vData
End Property

Public Property Get KepsEccentricity() As Single
  KepsEccentricity = mvarKepsEccentricity(mvarSatelliteIndex)
End Property

Public Property Let SatelliteElevation(ByVal vData As Long)
  mvarSatelliteElevation(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteElevation() As Long
  SatelliteElevation = mvarSatelliteElevation(mvarSatelliteIndex)
End Property

Public Property Let SatelliteAzimuth(ByVal vData As Long)
  mvarSatelliteAzimuth(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteAzimuth() As Long
  SatelliteAzimuth = mvarSatelliteAzimuth(mvarSatelliteIndex)
End Property

Public Property Let SatelliteRange(ByVal vData As Long)
  mvarSatelliteRange(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteRange() As Long
  SatelliteRange = mvarSatelliteRange(mvarSatelliteIndex)
End Property

Public Property Let SatelliteDesignator(ByVal vData As Variant)
  mvarSatelliteDesignator(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteDesignator() As Variant
  SatelliteDesignator = mvarSatelliteDesignator(mvarSatelliteIndex)
End Property

Public Property Let SatelliteName(ByVal vData As Variant)
  mvarSatelliteName(mvarSatelliteIndex) = vData
End Property

Public Property Get SatelliteName() As Variant
  SatelliteName = mvarSatelliteName(mvarSatelliteIndex)
End Property

Public Sub CalculateSatellitePosition()
  Dim SatMeanMotion As Variant
  Dim SatMeanMotionMinute As Variant
  Dim SatLinearDrag As Variant
  Dim DR As Variant
  Dim EA As Variant
  Dim C As Variant
  Dim S As Variant
  Dim DNOM As Variant
  Dim D As Variant
  Dim A As Variant
  Dim B As Variant
  Dim CW As Variant
  Dim SW As Variant
  Dim CQ As Variant
  Dim VELx As Variant
  Dim VELy As Variant
  Dim VELz As Variant
  Dim U As Variant
  Dim E As Variant
  Dim N As Variant
  Dim ElapsedTimeSinceEpoch As Variant

  mvarDataValid(mvarSatelliteIndex) = False
  
  mvarSatelliteDayNumber(mvarSatelliteIndex) = FNDAy(mvarDisplayYear(mvarSatelliteIndex), mvarDisplayMonth(mvarSatelliteIndex), mvarDisplayDay(mvarSatelliteIndex))
  mvarDisplayTimeRequired(mvarSatelliteIndex) = (mvarDisplayHour(mvarSatelliteIndex) + mvarDisplayMinute(mvarSatelliteIndex) / 60) / 24

  SetupSatellite

  ElapsedTimeSinceEpoch = (mvarSatelliteDayNumber(mvarSatelliteIndex) - SatEpochDayNumber) + (mvarDisplayTimeRequired(mvarSatelliteIndex) - mvarKepsYearEpochTimeFraction(mvarSatelliteIndex))
  SatMeanMotion = SatDragCoeff * ElapsedTimeSinceEpoch / 2
  SatMeanMotionMinute = 1 + 4 * SatMeanMotion
  SatLinearDrag = 1 - 7 * SatMeanMotion
  M = SatKepsMeanAnomoly + SatKepsMeanMotion * ElapsedTimeSinceEpoch * (1 - 3 * SatMeanMotion)
  DR = Int(M / (2 * PI))
  M = M - DR * 2 * PI
  mvarSatelliteOrbitNumber(mvarSatelliteIndex) = mvarKepsOrbitNumber(mvarSatelliteIndex) + DR

  EA = M
  Do
    C = Cos(EA)
    S = Sin(EA)
    DNOM = 1 - mvarKepsEccentricity(mvarSatelliteIndex) * C
    D = (EA - mvarKepsEccentricity(mvarSatelliteIndex) * S - M) / DNOM
    EA = EA - D
  Loop Until Abs(D) < 0.00001

  A = EarthRotationRateSeconds * SatMeanMotionMinute
  B = B0 * SatMeanMotionMinute
  RS(mvarSatelliteIndex) = A * DNOM
  Sx = A * (C - mvarKepsEccentricity(mvarSatelliteIndex))
  Vx = -A * S / DNOM * N0
  Sy = B * S: Vy = B * C / DNOM * N0
  AP = SatKepsArgOfPerigee + WD * ElapsedTimeSinceEpoch * SatLinearDrag
  CW = Cos(AP)
  SW = Sin(AP)
  RAAN = SatKepsRAAN + QD * ElapsedTimeSinceEpoch * SatLinearDrag
  CQ = Cos(RAAN)
  SQ = Sin(RAAN)
  CXx = CW * CQ - SW * CI * SQ
  CXy = -SW * CQ - CW * CI * SQ
  CXz = SI * SQ
  CYx = CW * SQ + SW * CI * CQ
  CYy = -SW * SQ + CW * CI * CQ
  CYz = -SI * CQ
  CZx = SW * SI
  CZy = CW * SI
  CZz = CI
  SatX = Sx * CXx + Sy * CXy
  ANTx = Ax * CXx + Ay * CXy + Az * CXz
  VELx = Vx * CXx + Vy * CXy
  SatY = Sx * CYx + Sy * CYy
  ANTy = Ax * CYx + Ay * CYy + Az * CYz
  VELy = Vx * CYx + Vy * CYy
  SatZ = Sx * CZx + Sy * CZy
  ANTz = Ax * CZx + Ay * CZy + Az * CZz
  VELz = Vx * CZx + Vy * CZy

  GHAA = GHAE + EarthRotationRateDay * ElapsedTimeSinceEpoch
  C = Cos(-GHAA)
  S = Sin(-GHAA)
  Sx = SatX * C - SatY * S
  Ax = ANTx * C - ANTy * S
  Vx = VELx * C - VELy * S
  Sy = SatX * S + SatY * C
  Ay = ANTx * S + ANTy * C
  Vy = VELx * S + VELy * C
  Sz = SatZ
  Az = ANTz
  Vz = VELz

  Rx = Sx - Ox
  Ry = Sy - Oy
  Rz = Sz - Oz
  mvarSatelliteRange(mvarSatelliteIndex) = Sqr(Rx ^ 2 + Ry ^ 2 + Rz ^ 2)
  Rx = Rx / mvarSatelliteRange(mvarSatelliteIndex)
  Ry = Ry / mvarSatelliteRange(mvarSatelliteIndex)
  Rz = Rz / mvarSatelliteRange(mvarSatelliteIndex)

  U = Rx * Ux + Ry * Uy + Rz * Uz
  E = Rx * Ex + Ry * Ey
  N = Rx * Nx + Ry * Ny + Rz * Nz

  mvarSatelliteAzimuth(mvarSatelliteIndex) = FNIntDEG(FNAtn(E, N))
  mvarSatelliteElevation(mvarSatelliteIndex) = FNIntDEG(FNASN(U))

  SQ = FNIntDEG(FNACS(-(Ax * Rx + Ay * Ry + Az * Rz)))

  mvarSatelliteLongitude(mvarSatelliteIndex) = FNIntDEG(FNAtn(Sy, Sx))
  mvarSatelliteLatitude(mvarSatelliteIndex) = FNIntDEG(FNASN(Sz / RS(mvarSatelliteIndex)))

  RR = (Vx - VOx) * Rx + (Vy - VOy) * Ry + Vz * Rz
  mvarDataValid(mvarSatelliteIndex) = True
  
End Sub

Sub CalculateALLPositions()

  Dim i As Integer
  Dim oldindex As Integer

  oldindex = mvarSatelliteIndex
  For i = 1 To 500
    If mvarSatelliteName(i) = "" Then Exit For
    mvarSatelliteIndex = i
    CalculateSatellitePosition
  Next i
  mvarSatelliteIndex = oldindex

End Sub
Private Function FNACS(X As Variant) As Variant

  FNACS = PI / 2 - Atn(X / Sqr(1 - X ^ 2))
End Function

Private Function FNASN(vValue As Variant) As Variant

  FNASN = Atn(vValue / Sqr(1 - vValue ^ 2))
End Function

Private Function FNAtn(Y As Variant, X As Variant) As Variant

  Dim Result As Variant

  If X <> 0 Then
    Result = Atn(Y / X)
  Else
    Result = PI / 2 * Sgn(Y)
  End If
  If X < 0 Then
    Result = Result + PI
  End If
  If Result < 0 Then
    Result = Result + 2 * PI
  End If
  FNAtn = Result
End Function
Private Function FNIntDEG(X As Variant) As Variant
  FNIntDEG = Int(X * 180 / PI)
End Function

Private Function FNRAD(X As Variant) As Variant
  FNRAD = X * PI / 180
End Function
Private Function FNDAy(Y As Variant, M As Variant, D As Variant) As Variant
  Dim TempY As Variant
  Dim TempM As Variant
  Dim TempD As Variant

  TempY = Y
  TempM = M
  TempD = D

  If TempM <= 2 Then TempY = TempY - 1: TempM = TempM + 12
  FNDAy = Int(TempY * MeanYear) + Int((TempM + 1) * 30.6) + TempD - 428
End Function

Private Sub Class_Initialize()

  PI = 4 * Atn(1)
  MeanYear = 365.25
  TropicalYear = 365.242197

  EarthRotationRate = 2 * PI / TropicalYear
  EarthRotationRateDay = 2 * PI + EarthRotationRate
  EarthRotationRateSeconds = EarthRotationRateDay / 86400
  RE = 6378.14
  FL = 1 / 298.257
  GravitationalConstant = 398600!
  ZonalCoeff = 0.00108263
  YG = 1990
  G0 = 99.4033

  'OLDRN = -99999
End Sub

Private Sub SetupSatellite()

  Dim XX As Variant
  Dim ZZ As Variant
  Dim D As Variant

  sLon = 180
  sLat = 0
  CL = Cos(mvarObserverLatitude)
  SL = Sin(mvarObserverLatitude)
  CO = Cos(mvarObserverLongitude)
  SO = Sin(mvarObserverLongitude)

  RP = RE * (1 - FL)
  XX = RE * RE
  ZZ = RP * RP

  D = Sqr(XX * CL * CL + ZZ * SL * SL)
  Rx = XX / D + mvarObserverHeight
  Rz = ZZ / D + mvarObserverHeight

  Ux = CL * CO
  Ex = -SO
  Nx = -SL * CO
  Uy = CL * SO
  Ey = CO
  Ny = -SL * SO
  Uz = SL
  Ez = 0
  Nz = CL
  Ox = Rx * Ux
  Oy = Rx * Uy
  Oz = Rz * Uz

  SatKepsRAAN = FNRAD(mvarKepsRAAN(mvarSatelliteIndex))
  SatKepsInclination = FNRAD(mvarKepsInclination(mvarSatelliteIndex))
  SatKepsArgOfPerigee = FNRAD(mvarKepsAOP(mvarSatelliteIndex))
  SatKepsMeanAnomoly = FNRAD(mvarKepsMeanAnomoly(mvarSatelliteIndex))
  SatKepsMeanMotion = mvarKepsMeanMotion(mvarSatelliteIndex) * 2 * PI

  M2 = mvarKepsDecayRate(mvarSatelliteIndex) * 2 * PI
  VOx = -Oy * EarthRotationRateSeconds
  VOy = Ox * EarthRotationRateSeconds

  SatEpochDayNumber = FNDAy(mvarSatEpochYear(mvarSatelliteIndex), 1, 0) + Int(mvarKepsYearEpochTime(mvarSatelliteIndex))
  mvarKepsYearEpochTimeFraction(mvarSatelliteIndex) = mvarKepsYearEpochTime(mvarSatelliteIndex) - Int(mvarKepsYearEpochTime(mvarSatelliteIndex))
  N0 = SatKepsMeanMotion / 86400
  EarthRotationRateSeconds = (GravitationalConstant / N0 / N0) ^ (1 / 3)
  B0 = EarthRotationRateSeconds * Sqr(1 - mvarKepsEccentricity(mvarSatelliteIndex) * mvarKepsEccentricity(mvarSatelliteIndex))
  SI = Sin(SatKepsInclination)
  CI = Cos(SatKepsInclination)
  PC = RE * EarthRotationRateSeconds / (B0 * B0)
  PC = 1.5 * ZonalCoeff * PC * PC * SatKepsMeanMotion
  QD = -PC * CI
  WD = PC * (5 * CI * CI - 1) / 2
  SatDragCoeff = -2 * M2 / SatKepsMeanMotion / 3
  TEG = (SatEpochDayNumber - FNDAy(YG%, 1, 0)) + mvarKepsYearEpochTimeFraction(mvarSatelliteIndex)
  GHAE = FNRAD(G0) + TEG * EarthRotationRateDay
  CO = Cos(FNRAD(sLon))
  SO = Sin(FNRAD(sLon))
  CL = Cos(FNRAD(sLat))
  SL = Sin(FNRAD(sLat))
  Ax = -CL * CO
  Ay = -CL * SO
  Az = -SL

End Sub

Sub DrawFootprints(pic As PictureBox)
  Dim clo As Variant
  Dim cla As Variant
  Dim TempX As Variant
  Dim TempY As Variant
  Dim TempZ As Variant
  Dim PointsToDraw As Integer
  Dim TwoPI As Variant
  Dim LineColour As Single
  Dim ScreenLON As Integer
  Dim ScreenLAT As Integer
  Dim SatX As Variant
  Dim SatY As Variant
  Dim SatZ As Variant
  Dim LastScreenLON As Integer
  Dim LastScreenLAT As Integer
  Dim ScrollCheck As Integer
  Dim srad As Variant
  Dim A As Variant
  Dim Counter As Integer

  PointsToDraw% = 40
  TwoPI = 2 * PI

  pic.Cls

  For i = 1 To 500
    Counter% = 0
    LastScreenLON = 0
    LastScreenLAT = 0
    If mvarSatelliteName(i) = "" Then Exit For
    ReDim FootprintLON(PointsToDraw%), FootprintLAT(PointsToDraw%)

    srad = FNACS(RE / RS(i))
    clo = FNRAD(mvarSatelliteLongitude(i))
    cla = FNRAD(mvarSatelliteLatitude(i))

    For A = 0 To TwoPI Step TwoPI / PointsToDraw%
      TempX = Cos(srad)
      TempY = Sin(srad) * Sin(A)
      TempZ = Sin(srad) * Cos(A)
      SatX = TempX * Cos(cla) - TempZ * Sin(cla)
      SatY = TempY
      SatZ = TempX * Sin(cla) + TempZ * Cos(cla)
      TempX = SatX * Cos(clo) - SatY * Sin(clo)
      TempY = SatX * Sin(clo) + SatY * Cos(clo)
      TempZ = SatZ
      FootprintLON(Counter%) = (FNAtn(TempY, TempX))
      FootprintLAT(Counter%) = (FNASN(TempZ))
      ScreenLON% = ConvertLONToScreen(FNIntDEG(FootprintLON(Counter%)))
      ScreenLAT% = ConvertLATToScreen(FNIntDEG(FootprintLAT(Counter%)))
      ScrollCheck% = (ScreenLON% - LastScreenLON%) * Sgn(ScreenLON% - LastScreenLON%)

      'Select Case ColourToDraw%
      '  Case 0, 1
      LineColour! = RGB(0, 255, 255)
      ' Case 2
      '   LineColour! = RGB(240, 240, 0)
      ' Case 3
      '   LineColour! = RGB(191, 191, 191)
      'End Select

      Select Case Counter%
        Case 0
          pic.PSet (ScreenLON%, ScreenLAT%)
        Case Else
          Select Case ScrollCheck%
            Case Is < 400
              pic.Line (LastScreenLON%, LastScreenLAT%)-(ScreenLON%, ScreenLAT%), LineColour!
            Case Is >= 400
              pic.PSet (ScreenLON%, ScreenLAT%)
          End Select
      End Select
      LastScreenLON% = ScreenLON%
      LastScreenLAT% = ScreenLAT%
      Counter% = Counter% + 1
    Next

    ScreenLON = Int(mvarSatelliteLongitude(i))
    ScreenLAT = Int(mvarSatelliteLatitude(i))
    ScreenLON% = ConvertLONToScreen(ScreenLON)
    ScreenLAT% = ConvertLATToScreen(ScreenLAT)

    pic.FillColor = RGB(255, 0, 0)
    pic.Circle (ScreenLON%, ScreenLAT%), 3, 0
    pic.ForeColor = RGB(255, 255, 255)
    TextOut pic.hdc, ScreenLON%, ScreenLAT%, mvarSatelliteName(i), Len(mvarSatelliteName(i))
  Next i
End Sub

Private Function ConvertLATToScreen(sLat As Integer) As Variant
  ConvertLATToScreen = gc000E% - sLat% * PixelsPerDegLat
End Function

Private Function ConvertLONToScreen(sLon As Integer) As Variant
  sLon% = 360 - sLon%
  MapCentre = 0
  Select Case MapCentre
Case 0:
    Select Case sLon%
      Case 0 To 180
        ConvertLONToScreen = gc000C% - sLon% * PixelsPerDegLon
      Case 181 To 360
        ConvertLONToScreen = gc000C% + (360 - sLon%) * PixelsPerDegLon
    End Select
Case 1
    ConvertLONToScreen = gc0020% - sLon% * PixelsPerDegLon
End Select
End Function
Sub DisplayAOS(AOSType As Integer)
Dim LastElevation As Variant
Dim GotAOS As Integer

LastElevation = mvarSatelliteElevation(mvarSatelliteIndex)
GotAOS% = 0
UpdateMap% = 0

Do
  Select Case AOSType
    Case 1
      DisplayMinute% = DisplayMinute% - 1
      TimeReqd = (DisplayHour% + DisplayMinute% / 60) / 24
      CalculateSatelliteVector
      CalculateRangeVector
      If SatelliteElevation < 0 And Sgn(SatelliteElevation) <> Sgn(LastElevation) Then
        GotAOS% = 1
      End If
      LastElevation = SatelliteElevation
      If DisplayMinute% < 0 Then
        DisplayMinute% = DisplayMinute% + 60
        DisplayHour% = DisplayHour% - 1
      End If
      If DisplayHour% < 0 Then
        DisplayHour% = DisplayHour% + 24
        DisplayDay% = DisplayDay% - 1
        SatDayNumber& = FNDAy(Y%, MN%, DisplayDay%)
      End If
    Case 2
      mvarDisplayMinute(mvarSatelliteIndex) = mvarDisplayMinute(mvarSatelliteIndex) + 1
      CalculateSatellitePosition
      If mvarSatelliteElevation(mvarSatelliteIndex) > 0 And Sgn(mvarSatelliteElevation(mvarSatelliteIndex)) <> Sgn(LastElevation) Then
        GotAOS% = 1
      End If
      LastElevation = mvarSatelliteElevation(mvarSatelliteIndex)
      If mvarDisplayMinute(mvarSatelliteIndex) > 59 Then
        mvarDisplayMinute(mvarSatelliteIndex) = mvarDisplayMinute(mvarSatelliteIndex) - 60
        mvarDisplayHour(mvarSatelliteIndex) = mvarDisplayHour(mvarSatelliteIndex) + 1
      End If
      If mvarDisplayHour(mvarSatelliteIndex) > 23 Then
        mvarDisplayHour(mvarSatelliteIndex) = mvarDisplayHour(mvarSatelliteIndex) - 24
        mvarDisplayDay(mvarSatelliteIndex) = mvarDisplayDay(mvarSatelliteIndex) + 1
        mvarSatelliteDayNumber(mvarSatelliteIndex) = FNDAy(Y%, MN%, DisplayDay%)
      End If
  End Select
Loop Until GotAOS% = 1
SatelliteLongitude% = SatLon
SatelliteLatitude% = SatLat
End Sub
Private mvarDisplayCentury(500) As Integer 'local copy
Private mvarDisplayYear(500) As Integer 'local copy
Private mvarDisplayMonth(500) As Integer 'local copy
Private mvarDisplayDay(500) As Integer 'local copy
Private mvarDisplayHour(500) As Integer 'local copy
Private mvarDisplayMinute(500) As Integer 'local copy
Private mvarSatelliteDayNumber(500) As Variant 'local copy
mvarDisplayTimeRequired
mvarSatelliteIndex
