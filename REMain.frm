VERSION 5.00
Begin VB.Form REMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renewable Energy Systems"
   ClientHeight    =   8850
   ClientLeft      =   4935
   ClientTop       =   4155
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "REMain.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   9720
   Begin VB.PictureBox REMainData 
      BackColor       =   &H80000009&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   9675
      TabIndex        =   1
      Top             =   2520
      Width           =   9735
   End
   Begin VB.PictureBox REMainHeader 
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu LoadTMY2 
         Caption         =   "Load TMY2"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Parameters 
      Caption         =   "Passive Solar"
      Enabled         =   0   'False
      Begin VB.Menu Collector 
         Caption         =   "Collectors"
      End
      Begin VB.Menu Systems 
         Caption         =   "System"
      End
      Begin VB.Menu Simulate 
         Caption         =   "Simulate"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu ASolar 
      Caption         =   "Active Solar"
   End
End
Attribute VB_Name = "REMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Collector_Click()
    Collectors.Show 1
    
End Sub
Private Sub Systems_Click()
    REMain.Enabled = False
    System.Show 1
End Sub
Private Sub Simulate_Click()
    If CDefine = 0 Then
        MsgBox ("Must Enter Collector Information")
    Else
        If SimSelect = 0 Then
            Call CalcQUseful
        ElseIf SimSelect = 1 Then
            Call CalcSimHouse
        ElseIf SimSelect = 2 Then
            Call CalcDHW
        ElseIf SimSelect = 3 Then
            Call CalcHouseWFan
        ElseIf SimSelect = 4 Then
            Call CalcHouseCombo
        Else
            MsgBox ("Must Define System")
        End If
    End If
        

End Sub

Private Sub CalcDHW()
            
    
    'FrTa = 0.725
    'FrUL = 0.83 * 10.76 * 1.055 * 1.8       '   kJ/hr-m2-C
    'numGlazings = 1
    'Ts = 22.22                              '  C
    'Tc1 = 22.22                             '  C
    'NumCollectors = 2
    'Area = NumCollectors * 24.5 * 1 / 10.76    '  m2
    'UAh = 1185                              '  kJ/hr-C
    'ms = 1500                               '  kg
    'cp = 4.19                               '  kJ/kg-C
    'Tg = 10                                 '  C
    'TsupW = 55                              '  C
    'e_hx = 0.8
    
    REMainHeader.Cls
    REMainHeader.Print "Calculate Domestic Hot Water"
    REMainHeader.Print "Weather Data: ", infile$
    REMainHeader.Print "Gamma: ", Gamma, "Beta: ", Beta, "RhoG: ", RhoG
    REMainHeader.Print "FrTa: ", Format$(FrTa, "0.00"), "FrUL: ", Format$(FrUL, "0.00")
    REMainHeader.Print "NumGlazings: ", numGlazings, "Collectors: ", NumCollectors
    REMainHeader.Print "Collector Area: ", Format$(Area, "0.00")
    REMainHeader.Print ""
             
    ReDim Ih_dy(365)
    ReDim Ih_mo(12)
    ReDim It_dy(365)
    ReDim It_mo(12)
    ReDim Ta_dy(365)
    ReDim Ta_mo(12)
    ReDim Qu_dy(365)
    ReDim Qu_mo(12)
    ReDim Qts_dy(365)
    ReDim Qts_mo(12)
    ReDim Qfs_dy(365)
    ReDim Qfs_mo(12)
    ReDim Qfsw_dy(365)
    ReDim Qfsw_mo(12)
    ReDim Qaux_dy(365)
    ReDim Qaux_mo(12)
    ReDim Qauxw_dy(365)
    ReDim Qauxw_mo(12)
    ReDim Qload_dy(365)
    ReDim Qload_mo(12)
    ReDim Qloadw_dy(365)
    ReDim Qloadw_mo(12)
    ReDim Ts_dy(365)
    ReDim Ts_mo(12)
    ReDim MWater(24)
    ReDim qU(8760)
     
                 
    '  Water usage array
    MWater(24) = 3              '  kg/hr
    MWater(1) = 1               '  kg/hr
    MWater(2) = 0.5             '  kg/hr
    MWater(3) = 0.5             '  kg/hr
    MWater(4) = 0.3             '  kg/hr
    MWater(5) = 1               '  kg/hr
    MWater(6) = 4               '  kg/hr
    MWater(7) = 17.5            '  kg/hr
    MWater(8) = 18              '  kg/hr
    MWater(9) = 22              '  kg/hr
    MWater(10) = 17.5           '  kg/hr
    MWater(11) = 14.5           '  kg/hr
    MWater(12) = 13             '  kg/hr
    MWater(13) = 12.5           '  kg/hr
    MWater(14) = 9              '  kg/hr
    MWater(15) = 8              '  kg/hr
    MWater(16) = 9              '  kg/hr
    MWater(17) = 10             '  kg/hr
    MWater(18) = 13             '  kg/hr
    MWater(19) = 18             '  kg/hr
    MWater(20) = 16             '  kg/hr
    MWater(21) = 13             '  kg/hr
    MWater(22) = 11             '  kg/hr
    MWater(23) = 10             '  kg/hr
        
    
    REMainData.Cls
    REMainData.Print "Month", "   It   ", "  Qts  ", "  Qfs  ", "  Qload  ", "  Qfsw  ", "  Qloadw  ", "  Ts  "
    REMainData.Print "     ", "kJ/dy-m2", "kJ/day ", "kJ/day ", " kJ/day  ", " kJ/day ", "  kJ/day  ", "  C   "
    REMainData.Print "-----", "--------", "-------", "-------", "---------", "--------", "----------", "------"
    
    For j = 1 To 8760
        
        
        b0 = -0.1
    
        kTa = 1 + b0 * ((1 / cosTheta(j)) - 1)

        If ((kTa < 0) Or (kTa > 1)) Then
            kTa = 0
        End If
        
        qU(j) = Area * (FrTa * kTa * iT(j) - FrUL * (Tc1 - Ta(j)))
               
        If qU(j) < 0 Then
            qU(j) = 0
        End If
        
        Tc1 = ((1 - e_hx) * Area * (FrTa * kTa * iT(j) + FrUL * Ta(j)) + (e_hx * MWater(hr(j)) * cp * Ts)) / (e_hx * MWater(hr(j)) * cp + (1 - e_hx) * Area * FrUL)
                
        If qU(j) > 0 And Ts <= 90 Then
            Qts = qU(j)
        Else
            Qts = 0
        End If
        
        Qfsw = MWater(hr(j)) * cp * (Ts - Tg)
        Qloadw = MWater(hr(j)) * cp * (TsupW - Tg)
        Qauxw = Qloadw - Qfsw
        
        If Quaxw < 0 Then
            Qauxw = 0
        End If
        
        '  Recalculate Ts
        Ts = Ts + (Qts - Qfsw) / (ms * cp)
                                              
        Call daysPerMonth(mo(j), days)
                                              
        '  Calculate Averages
        Ta_dy(n) = Ta_dy(n) + Ta(j) / 24
        Ta_mo(mo(j)) = Ta_mo(mo(j)) + Ta(j) / (24 * days)
        Ta_yr = Ta_yr + Ta(j) / 8760
        
        Ih_dy(n) = Ih_dy(n) + Ih(j)
        Ih_mo(mo(j)) = Ih_mo(mo(j)) + Ih(j) / days
        Ih_yr = Ih_yr + Ih(j) / 365
                
        It_dy(n) = It_dy(n) + iT(j)
        It_mo(mo(j)) = It_mo(mo(j)) + iT(j) / days
        It_yr = It_yr + iT(j) / 365
                
        Qu_dy(n) = Qu_dy(n) + qU(j)
        Qu_mo(mo(j)) = Qu_mo(mo(j)) + qU(j) / days
        Qu_yr = Qu_yr + qU(j) / 365
                
        Qts_dy(n) = Qts_dy(n) + Qts
        Qts_mo(mo(j)) = Qts_mo(mo(j)) + Qts / days
        Qts_yr = Qts_yr + Qts / 365
        
        Qfs_dy(n) = Qfs_dy(n) + Qfs
        Qfs_mo(mo(j)) = Qfs_mo(mo(j)) + Qfs / days
        Qfs_yr = Qfs_yr + Qfs / 365
        
        Qaux_dy(n) = Qaux_dy(n) + Qaux
        Qaux_mo(mo(j)) = Qaux_mo(mo(j)) + Qaux / days
        Qaux_yr = Qaux_yr + Qaux / 365
        
        Qload_dy(n) = Qload_dy(n) + Qload
        Qload_mo(mo(j)) = Qload_mo(mo(j)) + Qload / days
        Qload_yr = Qload_yr + Qload / 365
        
        Qfsw_dy(n) = Qfsw_dy(n) + Qfsw
        Qfsw_mo(mo(j)) = Qfsw_mo(mo(j)) + Qfsw / days
        Qfsw_yr = Qfsw_yr + Qfsw / 365
        
        Qauxw_dy(n) = Qauxw_dy(n) + Qauxw
        Qauxw_mo(mo(j)) = Qauxw_mo(mo(j)) + Qauxw / days
        Qauxw_yr = Qauxw_yr + Qauxw / 365
        
        Qloadw_dy(n) = Qloadw_dy(n) + Qloadw
        Qloadw_mo(mo(j)) = Qloadw_mo(mo(j)) + Qloadw / days
        Qloadw_yr = Qloadw_yr + Qloadw / 365
                
        Ts_dy(n) = Ts_dy(n) + Ts / 24
        Ts_mo(mo(j)) = Ts_mo(mo(j)) + Ts / (24 * days)
        Ts_yr = Ts_yr + Ts / 8760
                
    'If j = 647 Then Stop
    
    Next j
    
    Efficiency = Qts_yr / (Area * It_yr)
    SLR = (Qloadw_yr - Qauxw_yr) / Qloadw_yr
    
    For h = 1 To 12
        REMainData.Print h, Format$(It_mo(h), "0,0"), Format$(Qts_mo(h), "0,0"), Format$(Qfs_mo(h), "0,0"), Format$(Qload_mo(h), "0,0"), Format$(Qfsw_mo(h), "0,0"), Format$(Qloadw_mo(h), "0,0"), Format$(Ts_mo(h), "0.0")
    Next h
    
    REMainData.Print "-----", "----", "--------", "--------", "-------", "-------", "--------", "---------", "--"
    REMainData.Print "     ", Format$(It_yr, "0,0"), Format$(Qts_yr, "0,0"), Format$(Qfs_yr, "0"), Format$(Qload_yr, "0"), Format$(Qfsw_yr, "0"), Format$(Qloadw_yr, "0"), Format$(Ts_yr, "0.0")
    REMainData.Print ""
    REMainData.Print "Collector Efficiency: ", Format$(Efficiency, "0.00")
    REMainData.Print "Solar Load Ratio: ", Format$(SLR, "0.00")
End Sub


Private Sub CalcHouseCombo()
                
    
    'FrTa = 0.725
    'FrUL = 0.83 * 10.76 * 1.055 * 1.8 'kJ/hr-m2-C
    'numGlazings = 1
    'Ts = 48.9                          ' C
    'Tc1 = 22.22                         ' C
    'NumCollectors = 8
    'Area = NumCollectors * 24.5 * 1 / 10.76
    'UAh = 1185      ' kJ/hr-C
    'Tr = 22.22      ' C
    'Qelec = 2520    ' kJ/hr
    'mc = NumCollectors * 0.64 * 8.3237 * 1 / 2.2046 * 60
    'ml = mc
    'ms = 1500
    'cp = 4.19
    
    'Tg = 10             'C
    'TsupW = 55          'C
    'TsupH = 48.9        'C
    'e_hx = 0.8
    
    'e_loadhx = 0.25
    
    REMainHeader.Cls
    
    REMainHeader.Print "Calculate House with Fan and DHW"
    REMainHeader.Print "Weather Data: ", infile$
    REMainHeader.Print "Gamma: ", Gamma, "Beta: ", Beta, "RhoG: ", RhoG
    REMainHeader.Print "FrTa: ", FrTa, "FrUL: ", Format$(FrUL, "0.0")
    REMainHeader.Print "NumGlazings: ", numGlazings, "Collectors: ", NumCollectors
    REMainHeader.Print "Collector Area: ", Format$(Area, "0.00")
    REMainHeader.Print ""
    REMainHeader.Print "Tsupply: ", Ts
    REMainHeader.Print "Load HX Eff: ", e_loadhx
        
        '  Initialize
    
    ReDim Ih_dy(365)
    ReDim Ih_mo(12)
    ReDim It_dy(365)
    ReDim It_mo(12)
    ReDim Ta_dy(365)
    ReDim Ta_mo(12)
    ReDim Qu_dy(365)
    ReDim Qu_mo(12)
    ReDim Qts_dy(365)
    ReDim Qts_mo(12)
    ReDim Qfsh_dy(365)
    ReDim Qfsh_mo(12)
    ReDim Qfsw_dy(365)
    ReDim Qfsw_mo(12)
    ReDim Qaux_dy(365)
    ReDim Qaux_mo(12)
    ReDim Qauxw_dy(365)
    ReDim Qauxw_mo(12)
    ReDim Qloadh_dy(365)
    ReDim Qloadh_mo(12)
    ReDim Qloadw_dy(365)
    ReDim Qloadw_mo(12)
    ReDim Ts_dy(365)
    ReDim Ts_mo(12)
    ReDim MWater(24)
    ReDim qU(8760)
    
    
    Call setAvgArrays
               
    '  Water usage array
    MWater(24) = 3
    MWater(1) = 1
    MWater(2) = 0.5
    MWater(3) = 0.5
    MWater(4) = 0.3
    MWater(5) = 1
    MWater(6) = 4
    MWater(7) = 17.5
    MWater(8) = 18
    MWater(9) = 22
    MWater(10) = 17.5
    MWater(11) = 14.5
    MWater(12) = 13
    MWater(13) = 12.5
    MWater(14) = 9
    MWater(15) = 8
    MWater(16) = 9
    MWater(17) = 10
    MWater(18) = 13
    MWater(19) = 18
    MWater(20) = 16
    MWater(21) = 13
    MWater(22) = 11
    MWater(23) = 10
        
    REMainData.Cls
    REMainData.Print "Month", "   It   ", "  Qts  ", "  Qfsh  ", "  Qloadh ", "  Qfsw  ", "  Qloadw  ", "  Ts  "
    REMainData.Print "     ", "kJ/dy-m2", "kJ/day ", " kJ/day ", "  kJ/day ", " kJ/day ", "  kJ/day  ", "  C   "
    REMainData.Print "-----", "--------", "-------", "--------", "---------", "--------", "----------", "------"
    
    For j = 1 To 8760
        
        If numGlazings = 1 Then
            b0 = -0.1
        ElseIf numGlazings = 2 Then
            b0 = -0.17
        End If

        kTa = 1 + b0 * ((1 / cosTheta(j)) - 1)

        If ((kTa < 0) Or (kTa > 1)) Then
            kTa = 0
        End If
        
        qU(j) = Area * (FrTa * kTa * iT(j) - FrUL * (Tc1 - Ta(j)))
               
        If qU(j) < 0 Then
            qU(j) = 0
        End If
        
        Tc1 = ((1 - e_hx) * Area * (FrTa * kTa * iT(j) + FrUL * Ta(j)) + (e_hx * ml * cp * Ts)) / (e_hx * ml * cp + (1 - e_hx) * Area * FrUL)
            
        '  Calculate Q to thermal storage
        If qU(j) > 0 And Ts <= 90 Then
            Qts = qU(j)
        Else
            Qts = 0
        End If
            
        '  ============================================
        '  Calculate house heating load
        '  ============================================
            
        '  Calculate Q of the load
        QloadH = UAh * (Tr - Ta(j)) - Qelec
        If QloadH > 0 And (mo(j) >= 10 Or mo(j) <= 4) Then
            QloadH = QloadH
        Else
            QloadH = 0
        End If
            
        '  Calculate QfsH
        If QloadH > 0 And Ts > TsupH Then
            Qfsh = e_loadhx * 1274 * (Ts - Tr)
            If Qfsh < QloadH Then
                Qfsh = Qfsh
            Else
                Qfsh = QloadH
            End If
        Else
            Qfsh = 0
        End If
                    
        '  Calculate Qaux
        QauxH = QloadH - Qfsh
        
        '  ============================================
        '  Calculate DHW Heating Load
        '  ============================================
          
        Qfsw = MWater(hr(j)) * cp * (Ts - Tg)
        Qloadw = MWater(hr(j)) * cp * (TsupW - Tg)
        Qauxw = Qloadw - Qfsw
        
        If Quaxw < 0 Then
            Qauxw = 0
        End If
                
        Ts = Ts + (Qts - Qfsh - Qfsw) / (ms * cp)
        
        Call daysPerMonth(mo(j), days)
        
        '  Calculate Averages
        Ta_dy(n) = Ta_dy(n) + Ta(j) / 24
        Ta_mo(mo(j)) = Ta_mo(mo(j)) + Ta(j) / (24 * days)
        Ta_yr = Ta_yr + Ta(j) / 8760
        
        Ih_dy(n) = Ih_dy(n) + Ih(j)
        Ih_mo(mo(j)) = Ih_mo(mo(j)) + Ih(j) / days
        Ih_yr = Ih_yr + Ih(j) / 365
                
        It_dy(n) = It_dy(n) + iT(j)
        It_mo(mo(j)) = It_mo(mo(j)) + iT(j) / days
        It_yr = It_yr + iT(j) / 365
                
        Qu_dy(n) = Qu_dy(n) + qU(j)
        Qu_mo(mo(j)) = Qu_mo(mo(j)) + qU(j) / days
        Qu_yr = Qu_yr + qU(j) / 365
                
        Qts_dy(n) = Qts_dy(n) + Qts
        Qts_mo(mo(j)) = Qts_mo(mo(j)) + Qts / days
        Qts_yr = Qts_yr + Qts / 365
        
        Qfsh_dy(n) = Qfsh_dy(n) + Qfsh
        Qfsh_mo(mo(j)) = Qfsh_mo(mo(j)) + Qfsh / days
        Qfsh_yr = Qfsh_yr + Qfsh / 365
        
        Qfsw_dy(n) = Qfsw_dy(n) + Qfsw
        Qfsw_mo(mo(j)) = Qfsw_mo(mo(j)) + Qfsw / days
        Qfsw_yr = Qfsw_yr + Qfsw / 365
        
        Qauxw_dy(n) = Qauxw_dy(n) + Qauxw
        Qauxw_mo(mo(j)) = Qauxw_mo(mo(j)) + Qauxw / days
        Qauxw_yr = Qauxw_yr + Qauxw / 365
        
        Qloadh_dy(n) = Qloadh_dy(n) + QloadH
        Qloadh_mo(mo(j)) = Qloadh_mo(mo(j)) + QloadH / days
        Qloadh_yr = Qloadh_yr + QloadH / 365
        
        Qloadw_dy(n) = Qloadw_dy(n) + Qloadw
        Qloadw_mo(mo(j)) = Qloadw_mo(mo(j)) + Qloadw / days
        Qloadw_yr = Qloadw_yr + Qloadw / 365
                
        Ts_dy(n) = Ts_dy(n) + Ts / 24
        Ts_mo(mo(j)) = Ts_mo(mo(j)) + Ts / (24 * days)
        Ts_yr = Ts_yr + Ts / 8760
    
    Next j
    
    Efficiency = (Qfsh_yr + Qfsw_yr) / (Area * It_yr)
    SLR = (Qfsh_yr + Qfsw_yr) / (Qloadh_yr + Qloadw_yr)

    For h = 1 To 12
        REMainData.Print h, Format$(It_mo(h), "0,0"), Format$(Qts_mo(h), "0,0"), Format$(Qfsh_mo(h), "0,0"), Format$(Qloadh_mo(h), "0,0"), Format$(Qfsw_mo(h), "0,0"), Format$(Qloadw_mo(h), "0,0"), Format$(Ts_mo(h), "0.0")
    Next h
    
    REMainData.Print "-----", "--------", "-------", "--------", "---------", "--------", "----------", "------"
    REMainData.Print "Avg. ", Format$(It_yr, "0,0"), Format$(Qts_yr, "0,0"), Format$(Qfsh_yr, "0,0"), Format$(Qloadh_yr, "0,0"), Format$(Qfsw_yr, "0,0"), Format$(Qloadw_yr, "0,0"), Format$(Ts_yr, "0.0")
    REMainData.Print ""
    REMainData.Print "Collector Efficiency: ", Format$(Efficiency, "0.00")
    REMainData.Print "Solar Load Ratio: ", Format$(SLR, "0.00")

End Sub

Private Sub CalcHouseWFan()
                
    'FrTa = 0.725
    'FrUL = 0.83 * 10.76 * 1.055 * 1.8 'kJ/hr-m2-C
    'numGlazings = 1
    'Ts = 22.22                         ' C
    'Tc1 = 22.22                         ' C
    'NumCollectors = 8
    'Area = NumCollectors * 24.5 * 1 / 10.76
    'UAh = 1185      ' kJ/hr-C
    'Tr = 22.22      ' C
    'Qelec = 2520    ' kJ/hr
    'mc = NumCollectors * 0.64 * 8.3237 * 1 / 2.2046 * 60
    'ml = mc
    'ms = 1500
    'cp = 4.19
    
    'Tg = 10
    'TsupH = 48.9        '  C
    'e_hx = 0.8
    
    'e_loadhx = 0.25
    
    REMainHeader.Cls
    
    REMainHeader.Print "Calculate House with Fan"
    REMainHeader.Print "Weather Data: ", infile$
    REMainHeader.Print "Gamma: ", Gamma, "Beta: ", Beta, "RhoG: ", RhoG
    REMainHeader.Print "FrTa: ", FrTa, "FrUL: ", Format$(FrUL, "0.0")
    REMainHeader.Print "NumGlazings: ", numGlazings, "Collectors: ", NumCollectors
    REMainHeader.Print "Collector Area: ", Format$(Area, "0.00")
    REMainHeader.Print ""
    REMainHeader.Print "Tsupply: ", TsupH
    REMainHeader.Print "Load HX Eff: ", e_loadhx

       
    
    Call setAvgArrays
               
    '  Water usage array
    MWater(24) = 3
    MWater(1) = 1
    MWater(2) = 0.5
    MWater(3) = 0.5
    MWater(4) = 0.3
    MWater(5) = 1
    MWater(6) = 4
    MWater(7) = 17.5
    MWater(8) = 18
    MWater(9) = 22
    MWater(10) = 17.5
    MWater(11) = 14.5
    MWater(12) = 13
    MWater(13) = 12.5
    MWater(14) = 9
    MWater(15) = 8
    MWater(16) = 9
    MWater(17) = 10
    MWater(18) = 13
    MWater(19) = 18
    MWater(20) = 16
    MWater(21) = 13
    MWater(22) = 11
    MWater(23) = 10
        
    REMainData.Cls
    REMainData.Print "Month", "   It   ", "  Qts  ", "  Qfsh  ", " Qloadh ", "  Qfsw  ", "  Qloadw  ", "  Ts  "
    REMainData.Print "     ", "kJ/dy-m2", "kJ/day ", " kJ/day ", " kJ/day ", " kJ/day ", "  kJ/day  ", "  C   "
    REMainData.Print "-----", "--------", "-------", "--------", "--------", "--------", "----------", "------"
    
    For j = 1 To 8760
        
        If numGlazings = 1 Then
            b0 = -0.1
        ElseIf numGlazings = 2 Then
            b0 = -0.17
        End If

        kTa = 1 + b0 * ((1 / cosTheta(j)) - 1)

        If ((kTa < 0) Or (kTa > 1)) Then
            kTa = 0
        End If
        
        qU(j) = Area * (FrTa * kTa * iT(j) - FrUL * (Tc1 - Ta(j)))
               
        If qU(j) < 0 Then
            qU(j) = 0
        End If
        
        Tc1 = ((1 - e_hx) * Area * (FrTa * kTa * iT(j) + FrUL * Ta(j)) + (e_hx * ml * cp * Ts)) / (e_hx * ml * cp + (1 - e_hx) * Area * FrUL)
            
        '  Calculate Q to thermal storage
        If qU(j) > 0 And Ts <= 90 Then
            Qts = qU(j)
        Else
            Qts = 0
        End If
            
        '  ============================================
        '  Calculate house heating load
        '  ============================================
            
        '  Calculate Q of the load
        QloadH = UAh * (Tr - Ta(j)) - Qelec
        If QloadH > 0 And (mo(j) >= 10 Or mo(j) <= 4) Then
            QloadH = QloadH
        Else
            QloadH = 0
        End If
            
        '  Calculate QfsH
        If QloadH > 0 And Ts > TsupH Then
            Qfsh = e_loadhx * 1274 * (Ts - Tr)
            If Qfsh < QloadH Then
                Qfsh = Qfsh
            Else
                Qfsh = QloadH
            End If
        Else
            Qfsh = 0
        End If
                    
        '  Calculate Qaux
        QauxH = QloadH - Qfsh
                
        If QuaxH < 0 Then
            QauxH = 0
        End If
                
        Ts = Ts + (Qts - Qfsh) / (ms * cp)
        
        Call daysPerMonth(mo(j), days)
        
        '  Calculate Averages
        Ta_dy(n) = Ta_dy(n) + Ta(j) / 24
        Ta_mo(mo(j)) = Ta_mo(mo(j)) + Ta(j) / (24 * days)
        Ta_yr = Ta_yr + Ta(j) / 8760
        
        Ih_dy(n) = Ih_dy(n) + Ih(j)
        Ih_mo(mo(j)) = Ih_mo(mo(j)) + Ih(j) / days
        Ih_yr = Ih_yr + Ih(j) / 365
                
        It_dy(n) = It_dy(n) + iT(j)
        It_mo(mo(j)) = It_mo(mo(j)) + iT(j) / days
        It_yr = It_yr + iT(j) / 365
                
        Qu_dy(n) = Qu_dy(n) + qU(j)
        Qu_mo(mo(j)) = Qu_mo(mo(j)) + qU(j) / days
        Qu_yr = Qu_yr + qU(j) / 365
                
        Qts_dy(n) = Qts_dy(n) + Qts
        Qts_mo(mo(j)) = Qts_mo(mo(j)) + Qts / days
        Qts_yr = Qts_yr + Qts / 365
        
        Qfsh_dy(n) = Qfsh_dy(n) + Qfsh
        Qfsh_mo(mo(j)) = Qfsh_mo(mo(j)) + Qfsh / days
        Qfsh_yr = Qfsh_yr + Qfsh / 365
        
        Qfsw_dy(n) = Qfsw_dy(n) + Qfsw
        Qfsw_mo(mo(j)) = Qfsw_mo(mo(j)) + Qfsw / days
        Qfsw_yr = Qfsw_yr + Qfsw / 365
        
        Qauxw_dy(n) = Qauxw_dy(n) + Qauxw
        Qauxw_mo(mo(j)) = Qauxw_mo(mo(j)) + Qauxw / days
        Qauxw_yr = Qauxw_yr + Qauxw / 365
        
        Qloadh_dy(n) = Qloadh_dy(n) + QloadH
        Qloadh_mo(mo(j)) = Qloadh_mo(mo(j)) + QloadH / days
        Qloadh_yr = Qloadh_yr + QloadH / 365
        
        Qloadw_dy(n) = Qloadw_dy(n) + Qloadw
        Qloadw_mo(mo(j)) = Qloadw_mo(mo(j)) + Qloadw / days
        Qloadw_yr = Qloadw_yr + Qloadw / 365
                
        Ts_dy(n) = Ts_dy(n) + Ts / 24
        Ts_mo(mo(j)) = Ts_mo(mo(j)) + Ts / (24 * days)
        Ts_yr = Ts_yr + Ts / 8760
                
    'If j = 647 Then Stop
    
    Next j
    
    Efficiency = Qfsh_yr / (Area * It_yr)
    SLR = Qfsh_yr / Qloadh_yr

    For h = 1 To 12
        REMainData.Print h, Format$(It_mo(h), "0,0"), Format$(Qts_mo(h), "0,0"), Format$(Qfsh_mo(h), "0,0"), Format$(Qloadh_mo(h), "0,0"), Format$(Qfsw_mo(h), "0,0"), Format$(Qloadw_mo(h), "0,0"), Format$(Ts_mo(h), "0.0")
    Next h
    
    REMainData.Print "-----", "--------", "-------", "--------", "--------", "--------", "----------", "------"
    REMainData.Print "     ", Format$(It_yr, "0,0"), Format$(Qts_yr, "0,0"), Format$(Qfsh_yr, "0,0"), Format$(Qloadh_yr, "0,0"), Format$(Qfsw_yr, "0,0"), Format$(Qloadw_yr, "0,0"), Format$(Ts_yr, "0.0")
    REMainData.Print ""
    REMainData.Print "Collector Efficiency: ", Format$(Efficiency, "0.00")
    REMainData.Print "Solar Load Ratio: ", Format$(SLR, "0.00")
End Sub

Private Sub CalcSimHouse()
    
    '  Initialize
    
    'FrTa = 0.725
    'FrUL = 0.83 * 10.76 * 1.055 * 1.8 'kJ/hr-m2-C
    'numGlazings = 1
    'Ts = 22.22                            ' C
    'NumCollectors = 8
    'Area = NumCollectors * 24.5 * 1 / 10.76
    'UAh = 1185      ' kJ/hr-C
    'Tr = 22.22      ' C
    'Qelec = 2520    ' kJ/hr
    'mc = NumCollectors * 0.64 * 8.3237 * 1 / 2.2046 * 60
    'ml = mc
    'ms = 1500
    'cp = 4.19
    
    REMainHeader.Cls
    REMainHeader.Print "Calculate House Heating"
    REMainHeader.Print "Weather Data: ", infile$
    REMainHeader.Print "Gamma: ", Gamma, "Beta: ", Beta, "RhoG: ", RhoG
    REMainHeader.Print "FrTa: ", FrTa, "FrUL: ", Format$(FrUL, "0.0")
    REMainHeader.Print "NumGlazings: ", numGlazings, "Collectors: ", NumCollectors
    REMainHeader.Print "Collector Area: ", Format$(Area, "0.00")
    REMainHeader.Print ""
    REMainData.Cls
    REMainData.Print "Month", "   It   ", "  Qts  ", "  Qfs  ", "  Qload  ", "  Qfsw  ", "  Qloadw  ", "  Ts  "
    REMainData.Print "     ", "kJ/dy-m2", "kJ/day ", "kJ/day ", " kJ/day  ", " kJ/day ", "  kJ/day  ", "  C   "
    REMainData.Print "-----", "--------", "-------", "-------", "---------", "--------", "----------", "------"

    
    
    ReDim Ih_dy(365)
    ReDim Ih_mo(12)
    ReDim It_dy(365)
    ReDim It_mo(12)
    ReDim Ta_dy(365)
    ReDim Ta_mo(12)
    ReDim Qu_dy(365)
    ReDim Qu_mo(12)
    ReDim Qts_dy(365)
    ReDim Qts_mo(12)
    ReDim Qfs_dy(365)
    ReDim Qfs_mo(12)
    ReDim Qfsw_dy(365)
    ReDim Qfsw_mo(12)
    ReDim Qaux_dy(365)
    ReDim Qaux_mo(12)
    ReDim Qauxw_dy(365)
    ReDim Qauxw_mo(12)
    ReDim Qload_dy(365)
    ReDim Qload_mo(12)
    ReDim Qloadw_dy(365)
    ReDim Qloadw_mo(12)
    ReDim Ts_dy(365)
    ReDim Ts_mo(12)
    ReDim MWater(24)
    ReDim qU(8760)
            
    For j = 1 To 8760
               
        If numGlazings = 1 Then
            b0 = -0.1
        ElseIf numGlazings = 2 Then
            b0 = -0.17
        End If

        kTa = 1 + b0 * ((1 / cosTheta(j)) - 1)

        If ((kTa < 0) Or (kTa > 1)) Then
            kTa = 0
        End If
        
        qU(j) = Area * (FrTa * kTa * iT(j) - FrUL * (Ts - Ta(j)))
    
            
        '  Calculate Q to thermal storage
        If qU(j) > 0 And Ts <= 90 Then
            Qts = qU(j)
        Else
            Qts = 0
        End If
            
        '  Calculate Q of the load
        Qload = UAh * (Tr - Ta(j)) - Qelec
        If Qload > 0 And (mo(j) >= 10 Or mo(j) <= 4) Then
            Qload = Qload
        Else
            Qload = 0
        End If
            
        '  Calculate Qfs
        If Ts > Tr And Qload > 0 Then
            If (ml * cp * (Ts - Tr)) < Qload Then
                Qfs = ml * cp * (Ts - Tr)
            Else
                Qfs = Qload
            End If
        Else
            Qfs = 0
        End If
            
        '  Calculate Qaux
        Qaux = Qload - Qfs
        
        Ts = Ts + (Qts - Qfs) / (ms * cp)
                       
        Call daysPerMonth(mo(j), days)
                       
        Ta_dy(n) = Ta_dy(n) + Ta(j) / 24
        Ta_mo(mo(j)) = Ta_mo(mo(j)) + Ta(j) / (24 * days)
        Ta_yr = Ta_yr + Ta(j) / 8760
        
        Ih_dy(n) = Ih_dy(n) + Ih(j)
        Ih_mo(mo(j)) = Ih_mo(mo(j)) + Ih(j) / days
        Ih_yr = Ih_yr + Ih(j) / 365
                
        It_dy(n) = It_dy(n) + iT(j)
        It_mo(mo(j)) = It_mo(mo(j)) + iT(j) / days
        It_yr = It_yr + iT(j) / 365
                
        Qu_dy(n) = Qu_dy(n) + qU(j)
        Qu_mo(mo(j)) = Qu_mo(mo(j)) + qU(j) / days
        Qu_yr = Qu_yr + qU(j) / 365
                
        Qts_dy(n) = Qts_dy(n) + Qts
        Qts_mo(mo(j)) = Qts_mo(mo(j)) + Qts / days
        Qts_yr = Qts_yr + Qts / 365
        
        Qfs_dy(n) = Qfs_dy(n) + Qfs
        Qfs_mo(mo(j)) = Qfs_mo(mo(j)) + Qfs / days
        Qfs_yr = Qfs_yr + Qfs / 365
        
        Qaux_dy(n) = Qaux_dy(n) + Qaux
        Qaux_mo(mo(j)) = Qaux_mo(mo(j)) + Qaux / days
        Qaux_yr = Qaux_yr + Qaux / 365
        
        Qload_dy(n) = Qload_dy(n) + Qload
        Qload_mo(mo(j)) = Qload_mo(mo(j)) + Qload / days
        Qload_yr = Qload_yr + Qload / 365
             
        Qfsw_dy(n) = Qfsw_dy(n) + Qfsw
        Qfsw_mo(mo(j)) = Qfsw_mo(mo(j)) + Qfsw / days
        Qfsw_yr = Qfsw_yr + Qfsw / 365
                
        Qloadw_dy(n) = Qloadw_dy(n) + Qloadw
        Qloadw_mo(mo(j)) = Qloadw_mo(mo(j)) + Qloadw / days
        Qloadw_yr = Qloadw_yr + Qloadw / 365
             
        Ts_dy(n) = Ts_dy(n) + Ts / 24
        Ts_mo(mo(j)) = Ts_mo(mo(j)) + Ts / (24 * days)
        Ts_yr = Ts_yr + Ts / 8760
                
    'If j = 12 Then Stop
    
    Next j
    
    Efficiency = Qts_yr / (Area * It_yr)
    SLR = Qfs_yr / Qload_yr
    
    For h = 1 To 12
        REMainData.Print h, Format$(It_mo(h), "0"), Format$(Qts_mo(h), "0"), Format$(Qfs_mo(h), "0"), Format$(Qload_mo(h), "0"), Format$(Qfsw_mo(h), "0"), Format$(Qloadw_mo(h), "0"), Format$(Ts_mo(h), "0.0")
    Next h
    
    REMainData.Print "-----", "----", "--------", "--------", "-------", "-------", "--------", "---------", "--"
    REMainData.Print "     ", Format$(It_yr, "0"), Format$(Qts_yr, "0"), Format$(Qfs_yr, "0"), Format$(Qload_yr, "0"), Format$(Qfsw_yr, "0"), Format$(Qloadw_yr, "0"), Format$(Ts_yr, "0.0")
    REMainData.Print ""
    REMainData.Print "Collector Efficiency: ", Format$(Efficiency, "0.00")
    REMainData.Print "Solar Load Ratio: ", Format$(SLR, "0.00")
End Sub
Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
    SDefine = 0
    CDefine = 0
    
    '  On form load, define various values which we will use throughout the existance of this program
       
    '  PI
    pi = 3.141593

    '  Conversion -> Degrees to Radians
    dtor = 2 * pi / 360

    '  Conversion -> Radians to Degrees
    rtod = 360 / (2 * pi)

    '  Global Solar Constant
    GSC = 4920  'kJ/hr-m2
    
    cp = 4.19   'kJ/kg-C
End Sub

Private Sub LoadTMY2_Click()
    infile$ = "c:/Engineering Software/TMY2/Dayton.tm2"
    Open infile$ For Input As #1
    
    'Define Orientation
    'Gamma = 0
    'Beta = 40
    'RhoG = 0.2
    
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
    Line Input #1, I$
    timeZone = Val(Mid(I$, 34, 3))
    latitude = Val(Mid(I$, 40, 2)) + Val(Mid(I$, 43, 2)) / 60
    longitude = Val(Mid(I$, 48, 3)) + Val(Mid(I$, 52, 2)) / 60
    
    '  Process through all records of TMY2 file
    For j = 1 To 8760
        
        '  Reads line of TMY2 input
        Line Input #1, I$
    
        '  Extracts various fields of TMY2 data line
        mo(j) = Val(Mid(I$, 4, 2))          '  Extracts month
        dy(j) = Val(Mid(I$, 6, 2))          '  Extracts day
        hr(j) = Val(Mid(I$, 8, 2))          '  Extracts hour
        Ih(j) = Val(Mid(I$, 18, 4)) * 3.6   '  Extracts horizontal radiation (kJ/hour)
        Ta(j) = Val(Mid(I$, 68, 4)) / 10    '  Extracts outdoor air temperature (C)
        Tdp(j) = Val(Mid(I$, 74, 4)) / 10   '  Extracts dew point temperature (C)
        
        'Call SolDat(j)                      '  Calls SolDat() function, to calculate beam and diffuse radiation
        
    Next j
    
    Parameters.Enabled = True
    Collector.Enabled = True
    System.Enabled = True
    

End Sub
Private Sub CalcQUseful()
    
    'Tc1 = 30    '  C
       
    REMainHeader.Cls
    REMainHeader.Print "Calculate Qu"
    REMainHeader.Print "Weather Data: ", infile$
    REMainHeader.Print "Gamma: ", Gamma, "Beta: ", Beta, "RhoG: ", RhoG
    REMainHeader.Print "FrTa: ", Format$(FrTa, "0.0"), "FrUL: ", Format$(FrUL, "0.0")
    REMainHeader.Print "NumGlazings: ", numGlazings, "Collectors: ", NumCollectors
    REMainHeader.Print "Collector Area: ", Format$(Area, "0.00")
    REMainHeader.Print ""
    REMainData.Cls
    REMainData.Print "Month", "Tavg", "  I_Avg ", " It_Avg ", "Qu_Avg"
    REMainData.Print "     ", "(C) ", "kJ/dy-m2", "kJ/dy-m2", "kJ/day"
    REMainData.Print "-----", "----", "--------", "--------", "------"
    
    '  Initialize
    '  Call setDaysperMonth
    '  Call setAvgArrays

    ReDim Ih_dy(365)
    ReDim Ih_mo(12)
    ReDim It_dy(365)
    ReDim It_mo(12)
    ReDim Ta_dy(365)
    ReDim Ta_mo(12)
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
    ReDim qU(8760)
           
    For j = 1 To 8760
        
        b0 = -0.1
        kTa = 1 + b0 * ((1 / cosTheta(j)) - 1)

        If ((kTa < 0) Or (kTa > 1)) Then
            kTa = 0
        End If
        
        qU(j) = Area * (FrTa * kTa * iT(j) - FrUL * (Tc1 - Ta(j)))
        
        If (qU(j) < 0) Then
            qU(j) = 0
        End If
        
        Call daysPerMonth(mo(j), days)
        
        Ta_mo(mo(j)) = Ta_mo(mo(j)) + Ta(j) / (24 * days)
        Ta_yr = Ta_yr + Ta(j) / 8760
        
        Ih_mo(mo(j)) = Ih_mo(mo(j)) + Ih(j) / days
        Ih_yr = Ih_yr + Ih(j) / 365
                
        It_mo(mo(j)) = It_mo(mo(j)) + iT(j) / days
        It_yr = It_yr + iT(j) / 365
                
        Qu_mo(mo(j)) = Qu_mo(mo(j)) + qU(j) / days
        Qu_yr = Qu_yr + qU(j) / 365
               
    Next j
    
    Efficiency = Qu_yr / It_yr
    
    For h = 1 To 12
        REMainData.Print h, Format$(Ta_mo(h), "0,0.0"), Format$(Ih_mo(h), "0"), Format$(It_mo(h), "0"), Format$(Qu_mo(h), "0")
    Next h
    
    REMainData.Print "-----", "----", "--------", "--------", "------"
    REMainData.Print "     ", Format$(Ta_yr, "0.0"), Format$(Ih_yr, "0"), Format$(It_yr, "0"), Format$(Qu_yr, "0")
    REMainData.Print ""
    REMainData.Print "Collector Efficiency: ", Format$(Efficiency, "0.00")
End Sub


