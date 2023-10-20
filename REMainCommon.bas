Attribute VB_Name = "REMainCommon"
Option Explicit
Public Sub CalcSHGCB(angle, coeff1, coeff2, coeff3, coeff4, SHGCB_value)
    If (angle = 0) Then
        SHGCB_value = coeff1 + coeff2 + coeff3 + coeff4
    
    Else
        SHGCB_value = coeff1 + coeff2 * angle + coeff3 * (angle) ^ 2 + coeff4 * (angle) ^ 3
    End If

End Sub

Public Sub SolDat(j)

    Dim n
    Dim ndeg
    Dim eqnTime
    Dim declination
    Dim hrsol
    Dim wsol
    Dim cosTheta1, cosTheta2, cosTheta3, cosTheta4, cosTheta5
    Dim Rb
    Dim kT
    Dim erb
    
    Dim Y
    Dim c0
    Dim c1
    Dim c2
    Dim c3
    Dim weW
    Dim gS, gW
    Dim cW
    
    Dim iDiffuse, iBeam, I0, iTiltedBeam, iTiltedBeamO
    Dim iTotalDif, iTotalDifO, iTotalDifCS, iTotalDifIso, iTotalDifHor, iTotalDifRef
    Dim aI
    
    Dim SHGCDif, SHGCBnorm, SHGCBeam, SHGCBtheta
    
    Dim dt
    Dim k
    
    Dim Tewp, Tfp, Tcp, Tiwp
    
    Dim Tianum, Tiaden
    Dim Tset
    
    
    
    '  ===================================================== '
    '  STEP 1 - Find wSol, hour angle of solar time
    '  ===================================================== '
    
    '  Calculate day number, n
    n = 1 + Int((j - 1) / 24)
        
    '  Calculate day in degrees
    ndeg = (n - 1) * 360 / 365
    
    '  Calculate equation of time in degrees
    eqnTime = 229.2 * (0.000075 + 0.001868 * Cos(ndeg * dtor) - 0.032077 * Sin(ndeg * dtor) - 0.014615 * Cos(2 * ndeg * dtor) - 0.04089 * Sin(2 * ndeg * dtor))
    
    '  Calculate declination in degrees
    declination = 23.45 * Sin(360 * (284 + n) / 365 * dtor)
    
    '  Calculate solar hour
    hrsol = hr(j) + (4 * ((timeZone * -15) - longitude) + eqnTime) / 60
    
    '  Calculate solar hour angle
    wsol = ((hrsol - 12) * 15) - 7.5
    
    '  ===================================================== '
    '  STEP 2 - Find CosTheta and CosThetaZ
    '  ===================================================== '
    
    '  Calculate components of cosTheta
    cosTheta1 = Sin(declination * dtor) * Sin(latitude * dtor) * Cos(Beta * dtor)
    cosTheta2 = Sin(declination * dtor) * Cos(latitude * dtor) * Sin(Beta * dtor) * Cos(Gamma * dtor)
    cosTheta3 = Cos(declination * dtor) * Cos(latitude * dtor) * Cos(Beta * dtor) * Cos(wsol * dtor)
    cosTheta4 = Cos(declination * dtor) * Sin(latitude * dtor) * Sin(Beta * dtor) * Cos(Gamma * dtor) * Cos(wsol * dtor)
    cosTheta5 = Cos(declination * dtor) * Sin(Beta * dtor) * Sin(Gamma * dtor) * Sin(wsol * dtor)
        
    '  Calculate total cosTheta
    cosTheta(j) = cosTheta1 - cosTheta2 + cosTheta3 + cosTheta4 + cosTheta5
    
    '  Calculate cosThetaZ, angle between beam and normal which is horizontal to collector surface
    cosThetaZ(j) = Cos(latitude * dtor) * Cos(declination * dtor) * Cos(wsol * dtor) + Sin(latitude * dtor) * Sin(declination * dtor)
    
    '  ===================================================== '
    '  STEP 3 - Find Ibeam and Idiffuse using Erb's Relation
    '  ===================================================== '
    
    '  Calculate ratio between beam radiation on collector/beam radiation on horizontal surface
    Rb = cosTheta(j) / cosThetaZ(j)
    
    '  Perform bounds check on angles and angle
    If cosTheta(j) < 0 Or cosThetaZ(j) < 0 Then Rb = 0
    If Rb > 2.8 Then Rb = 2.8
    
    '  Calculate mean radiation at edge of atmosphere
    '  parrallel to earth's surface
    I0 = GSC * (1 + 0.033 * Cos((360 * n / 365) * dtor)) * cosThetaZ(j)
    
    '  Bounds check
    If I0 < 0 Then
        I0 = 0
    End If
    
    '  Calculate radiation components because is daytime
    If Ih(j) > 0 And I0 > 0 And I0 > Ih(j) Then
                 
        '  Calculate hourly clearness index, kT
        kT = Ih(j) / I0
    
        '  Calculate Erb's relation based on hourly clearness index
        If (kT <= 0.22) Then
            erb = 1 - 0.09 * kT
        ElseIf (kT <= 0.8) Or (kT > 0.22) Then
            erb = 0.9511 - 0.1604 * kT + 4.388 * (kT ^ 2) - 16.638 * (kT ^ 3) + 12.336 * (kT ^ 4)
        Else
            erb = 0.165
        End If
        
        '  Calculate diffuse radiation component
        iDiffuse = erb * Ih(j)
        
        '  Calculate beam radiation component
        iBeam = Ih(j) - iDiffuse
        
        '  ===================================================== '
        '  STEP 4 - Find ItiltedBeam, beam radiation on tilted surface
        '  ===================================================== '
                
        '  Calculate radiation on tilted surface
        iTiltedBeam = iBeam * Rb
        
        '  ===================================================== '
        '  STEP 5 - Find ItiltedDiffuse, diffuse radiation on tilted surface
        '  ===================================================== '
        
        '  ===================================================== '
        '  STEP PROTRUSION - Addition of protrusion specific calculations for passive solar
        '  ===================================================== '
        
        '  Calculate overhang shading coefficient
       ' If pO > 0 And hCO > 0 Then
       '     Y = pO / Tan(invcos(cosThetaZ(j)))
       '     c0 = (gapO + hCO - Y) / hCO
       '     If c0 < 0 Then c0 = 0
       '     If c0 > 1 Then c0 = 1
       ' Else
       '     c0 = 1
       ' End If
        
        '  Calculate Azimuth angle of solar to a wall at any orientation
       ' If Abs(Tan(declination * dtor) / Tan(latitude * dtor)) > 1 Then
            c1 = 1
       ' Else
        '    weW = invcos(Tan(declination * dtor) / Tan(latitude * dtor)) * rtod
        '    If Abs(wsol) < weW Then
        '        c1 = 1
        '    Else
        '        c1 = -1
        '    End If
        'End If
        
        'If latitude * (latitude - declination) >= 0 Then
        '    c2 = 1
        'Else
        '    c2 = -1
        'End If
        
        'If wsol >= 0 Then
        '    c3 = 1
        'Else
        '    c3 = -1
        'End If
        
        'gS = c1 * c2 * invsin(Sin(wsol * dtor) * Cos(declination * dtor) / Sin(invcos(cosThetaZ(j)))) * rtod + c3 * (1 - c1 * c2) * 90
        
        'gW = gS - Gamma
        
        '  Calculate vertical fin shading coefficient
        'If pW > 0 And wC > 0 Then
        '    Y = pW * Tan(Abs(gW) * dtor)
        '    If wC <> 0 Then cW = (gapW + wC - Y) / wC
        '    If cW <= 0 Then cW = 0
        '    If cW > 1 Then cW = 1
        'Else
        '    cW = 1
        'End If
            
        'If Abs(gW) >= 90 Then cW = 0
            '  Calculate ansiotropic index
            aI = iBeam / I0
        
            '  Total diffuse radiation from circumsolar radiation
            iTotalDifCS = iDiffuse * Rb * aI
        
            '  Total diffuse radiation spread evenly over sky
            iTotalDifIso = iDiffuse * ((1 + Cos(Beta * dtor)) / 2) * (1 - aI)
            
            '  Total diffuse radiation from area near horizon
            iTotalDifHor = iTotalDifIso * ((iBeam / Ih(j)) ^ 0.5) * (Sin((Beta * dtor) / 2)) ^ 3
            
            '  Total diffuse radiation from surfaces seen by collector
            iTotalDifRef = Ih(j) * RhoG * ((1 - Cos(Beta * dtor)) / 2)

            '  Total diffuse radition
            iTotalDif = iTotalDifCS + iTotalDifIso + iTotalDifHor + iTotalDifRef
         '   iTotalDifO = iTotalDifCS * cW * c0 + iTotalDifIso + iTotalDifHor + iTotalDifRef
            
            '  ===================================================== '
            '  STEP 6 - Find iT
            '  ===================================================== '
        
            '  Calculate iTiltedBeamO
        '    iTiltedBeamO = iTiltedBeam * cW * c0
                
            '  Calculate total radiation according to Hay, Davies, Klucher, Reindl model
            iT(j) = iTiltedBeam + iTotalDif
        '    iTO(j) = iTiltedBeamO + iTotalDifO
        
        Else    '  Night time
            
        '    cW = 0
        '    c0 = 0
            
            iTiltedBeam = 0
            iTotalDif = 0
            
        '    iTiltedBeamO = iTiltedBeam * cW * c0
        '    iTotalDifO = 0
                        
            iT(j) = 0   ' No solar radiation at night
        '    iTO(j) = iTiltedBeamO + iTotalDifO
                   
        End If
        
        
    
End Sub

Private Sub LoadTMY2Data()
    infile$ = "c:/Engineering Software/TMY2/dayton.tm2"
    Open infile$ For Input As #1
    
    '  Define Orientation
    Gamma = 0
    Beta = 0
    RhoG = 0.2
    
    pO = 0
    gapO = 1
    hCO = 4
    pW = 0
    gapW = 1
    wC = 4
        
    ' Dimension
    ReDim mo(8760)
    ReDim dy(8760)
    ReDim hr(8760)
    ReDim Ih(8760)
    ReDim Ta(8760)
    ReDim Tdp(8760)
    ReDim cosTheta(8760)
    ReDim cosThetaZ(8760)
    ReDim iT(8760)
    ReDim iTO(8760)
    
    'process header line
    Line Input #1, i$
    timeZone = Val(Mid(i$, 34, 3))
    latitude = Val(Mid(i$, 40, 2)) + Val(Mid(i$, 43, 2)) / 60
    longitude = Val(Mid(i$, 48, 3)) + Val(Mid(i$, 52, 2)) / 60
    
    '  Process through all records of TMY2 file
    For j = 1 To 8760
        
        '  Reads line of TMY2 input
        Line Input #1, i$
    
        '  Extracts various fields of TMY2 data line
        mo(j) = Val(Mid(i$, 4, 2))          '  Extracts month
        dy(j) = Val(Mid(i$, 6, 2))          '  Extracts day
        hr(j) = Val(Mid(i$, 8, 2))          '  Extracts hour
        Ih(j) = Val(Mid(i$, 18, 4)) * 3.6   '  Extracts horizontal radiation (kJ/hour)
        Ta(j) = Val(Mid(i$, 68, 4)) / 10    '  Extracts outdoor air temperature (C)
        Tdp(j) = Val(Mid(i$, 74, 4)) / 10   '  Extracts dew point temperature (C)
        
        Call SolDat(j)                      '  Calls SolDat() function, to calculate beam and diffuse radiation
        
    Next j
    
    Calculate.Enabled = True

End Sub

Public Sub daysPerMonth(X As Variant, Y As Variant)
    If X = 1 Then Y = 31    ' January
    If X = 2 Then Y = 28    ' February
    If X = 3 Then Y = 31    ' March
    If X = 4 Then Y = 30    ' April
    If X = 5 Then Y = 31    ' May
    If X = 6 Then Y = 30    ' June
    If X = 7 Then Y = 31    ' July
    If X = 8 Then Y = 31    ' August
    If X = 9 Then Y = 30    ' September
    If X = 10 Then Y = 31   ' October
    If X = 11 Then Y = 30   ' November
    If X = 12 Then Y = 31   ' December
End Sub

Public Sub setAvgArrays()
    
    ReDim Ih_dy(365)
    ReDim Ih_mo(12)
    ReDim It_dy(365)
    ReDim It_mo(12)
    ReDim Ta_dy(365)
    ReDim Ta_mo(12)
    ReDim daysPerMonth(12)
    ReDim Qu_dy(365)
    ReDim Qu_mo(12)
    ReDim Qts_dy(365)
    ReDim Qts_mo(12)
    ReDim Qfs_dy(365)
    ReDim Qfs_mo(12)
    ReDim Qfsw_dy(365)
    ReDim Qfsw_mo(12)
    ReDim Qauxw_dy(365)
    ReDim Qauxw_mo(12)
    ReDim Qload_dy(365)
    ReDim Qload_mo(12)
    ReDim Qloadw_dy(365)
    ReDim Qloadw_mo(12)
    ReDim Ts_dy(365)
    ReDim Ts_mo(12)
    ReDim MWater(24)
End Sub

Public Function invcos(X)
    invcos = Atn(-X / Sqr(-X * X + 1)) + 1.5708
End Function

Public Function invsin(X)
    invsin = Atn(X / Sqr(-X * X + 1))
End Function
