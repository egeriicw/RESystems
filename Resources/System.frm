VERSION 5.00
Begin VB.Form System 
   Caption         =   "System Parameters"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "**Note**"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   45
      Top             =   5280
      Width           =   3015
      Begin VB.Label Label1 
         Caption         =   $"System.frx":0000
         Height          =   1095
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton SystemCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   44
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton SystemOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   43
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame FHXFrame 
      Caption         =   "Fan Heat Exchanger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   34
      Top             =   4560
      Width           =   3855
      Begin VB.TextBox FHXText 
         Height          =   285
         Left            =   3120
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.Label FHXLabel 
         Caption         =   "Effectiveness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame CHXFrame 
      Caption         =   "Collector Heat Exchanger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   31
      Top             =   4200
      Width           =   3015
      Begin VB.TextBox CHXText 
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   480
         Width           =   615
      End
      Begin VB.Label CHXLabel 
         Caption         =   "Effectiveness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.TextBox HText5 
      Height          =   285
      Left            =   6480
      TabIndex        =   18
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox HText4 
      Height          =   285
      Left            =   6480
      TabIndex        =   17
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox HText3 
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.Frame WFrame 
      Caption         =   "Load - Water"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3360
      TabIndex        =   12
      Top             =   5520
      Width           =   3855
      Begin VB.TextBox WText2 
         Height          =   285
         Left            =   3120
         TabIndex        =   41
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox WText1 
         Height          =   285
         Left            =   3120
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.Label WLabel2 
         Caption         =   "Supply Water Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label WLabel1 
         Caption         =   "Ground Water Temperature (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame HFrame 
      Caption         =   "Load - House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton HCheck2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton HCheck1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox HText8 
         Height          =   285
         Left            =   3120
         TabIndex        =   39
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox HText7 
         Height          =   285
         Left            =   3120
         TabIndex        =   29
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox HText2 
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox HText6 
         Height          =   285
         Left            =   3120
         TabIndex        =   24
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox HText1 
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label HLabel9 
         Caption         =   "Supply Air Temperature (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label HLabel8 
         Caption         =   "House Set Point Temperature (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label HLabel7 
         Caption         =   "Electrical Load (W)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label HLabel6 
         Caption         =   "UA Infiltration (hr-ft2-F/Btu)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label HLabel5 
         Caption         =   "UA Ceiling (hr-ft2-F/Btu)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label HLabel4 
         Caption         =   "UA Windows (hr-ft2-F/Btu)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label HLabel3 
         Caption         =   "UA Walls (hr-ft2-F/Btu)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label HLabel2 
         Caption         =   "Individual UA Values"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label HLabel1 
         Caption         =   "Total UA Value (hr-ft2-F/Btu)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame CFrame 
      Caption         =   "Collector Fluid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   3015
      Begin VB.TextBox CText2 
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox CText1 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label CLabel2 
         Caption         =   "Flow Rate (gpm)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label CLabel1 
         Caption         =   "Temperature of Fluid Entering Collector (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Model Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox ModelCombo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame TSFrame 
      Caption         =   "Thermal Storage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
      Begin VB.CheckBox Check1 
         Caption         =   "Include Stratification"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TSText2 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TSText1 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label TSLabel2 
         Caption         =   "Volume (m3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label TSLabel1 
         Caption         =   "Initial Temperature (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label4 
      Caption         =   "UAWalls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CFrame.Enabled = False
    CText1.Enabled = False
    CText2.Enabled = False
    CLabel1.Enabled = False
    CLabel2.Enabled = False
    
    TSFrame.Enabled = False
    TSLabel1.Enabled = False
    TSLabel2.Enabled = False
    TSText1.Enabled = False
    TSText2.Enabled = False
    
    HFrame.Enabled = False
    HLabel1.Enabled = False
    HLabel2.Enabled = False
    HLabel3.Enabled = False
    HLabel4.Enabled = False
    HLabel5.Enabled = False
    HLabel6.Enabled = False
    HLabel7.Enabled = False
    HLabel8.Enabled = False
    HLabel9.Enabled = False
    HText1.Enabled = False
    HText2.Enabled = False
    HText3.Enabled = False
    HText4.Enabled = False
    HText5.Enabled = False
    HText6.Enabled = False
    HText7.Enabled = False
    HText8.Enabled = False
    HCheck1.Enabled = False
    HCheck2.Enabled = False
    
    FHXFrame.Enabled = False
    FHXLabel.Enabled = False
    FHXText.Enabled = False
    
    WFrame.Enabled = False
    WLabel1.Enabled = False
    WLabel2.Enabled = False
    WText1.Enabled = False
    WText2.Enabled = False
         
    CHXFrame.Enabled = False
    CHXLabel.Enabled = False
    CHXText.Enabled = False
       
    ModelCombo.AddItem ("Q Useful")
    ModelCombo.AddItem ("House Heating")
    ModelCombo.AddItem ("Domestic Hot Water")
    ModelCombo.AddItem ("House Heating with Fan Coil")
    ModelCombo.AddItem ("House Heating w/ Fan and DHW")
End Sub



Private Sub HCheck1_Click()
    If HCheck1.Value = True Then
        HCheck1.Value = True
        
        HLabel1.Enabled = True
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HText1.Enabled = True
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        
    End If
    
End Sub

Private Sub HCheck2_Click()
    If HCheck2.Value = True Then
        
        HCheck2.Value = True
    
        
        HLabel1.Enabled = False
        HLabel2.Enabled = True
        HLabel3.Enabled = True
        HLabel4.Enabled = True
        HLabel5.Enabled = True
        HLabel6.Enabled = True
        HText1.Enabled = False
        HText2.Enabled = True
        HText3.Enabled = True
        HText4.Enabled = True
        HText5.Enabled = True
        
    End If
End Sub

Private Sub ModelCombo_Click()
    
    Dim temp
    
    temp = ModelCombo.ListIndex
    
    If temp = 0 Then
        CFrame.Enabled = True
        CText1.Enabled = True
        CText2.Enabled = False
        CLabel1.Enabled = True
        CLabel2.Enabled = False
    
        TSFrame.Enabled = False
        TSLabel1.Enabled = False
        TSLabel2.Enabled = False
        TSText1.Enabled = False
        TSText2.Enabled = False
    
        HFrame.Enabled = False
        HLabel1.Enabled = False
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HLabel7.Enabled = False
        HLabel8.Enabled = False
        HLabel9.Enabled = False
        HText1.Enabled = False
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        HText6.Enabled = False
        HText7.Enabled = False
        HText8.Enabled = False
        HCheck1.Enabled = False
        HCheck2.Enabled = False
    
        FHXFrame.Enabled = False
        FHXLabel.Enabled = False
        FHXText.Enabled = False
    
        WFrame.Enabled = False
        WLabel1.Enabled = False
        WLabel2.Enabled = False
        WText1.Enabled = False
        WText2.Enabled = False
         
        CHXFrame.Enabled = False
        CHXLabel.Enabled = False
        CHXText.Enabled = False
    ElseIf temp = 1 Then
        CFrame.Enabled = False
        CText1.Enabled = False
        CText2.Enabled = False
        CLabel1.Enabled = False
        CLabel2.Enabled = False
    
        TSFrame.Enabled = False
        TSLabel1.Enabled = False
        TSLabel2.Enabled = False
        TSText1.Enabled = False
        TSText2.Enabled = False
    
        HFrame.Enabled = False
        HLabel1.Enabled = False
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HLabel7.Enabled = False
        HLabel8.Enabled = False
        HLabel9.Enabled = False
        HText1.Enabled = False
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        HText6.Enabled = False
        HText7.Enabled = False
        HText8.Enabled = False
        HCheck1.Enabled = False
        HCheck2.Enabled = False
    
        FHXFrame.Enabled = False
        FHXLabel.Enabled = False
        FHXText.Enabled = False
    
        WFrame.Enabled = False
        WLabel1.Enabled = False
        WLabel2.Enabled = False
        WText1.Enabled = False
        WText2.Enabled = False
         
        CHXFrame.Enabled = False
        CHXLabel.Enabled = False
        CHXText.Enabled = False
    ElseIf temp = 2 Then
        CFrame.Enabled = False
        CText1.Enabled = False
        CText2.Enabled = False
        CLabel1.Enabled = False
        CLabel2.Enabled = False
    
        TSFrame.Enabled = False
        TSLabel1.Enabled = False
        TSLabel2.Enabled = False
        TSText1.Enabled = False
        TSText2.Enabled = False
    
        HFrame.Enabled = False
        HLabel1.Enabled = False
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HLabel7.Enabled = False
        HLabel8.Enabled = False
        HLabel9.Enabled = False
        HText1.Enabled = False
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        HText6.Enabled = False
        HText7.Enabled = False
        HText8.Enabled = False
        HCheck1.Enabled = False
        HCheck2.Enabled = False
    
        FHXFrame.Enabled = False
        FHXLabel.Enabled = False
        FHXText.Enabled = False
    
        WFrame.Enabled = False
        WLabel1.Enabled = False
        WLabel2.Enabled = False
        WText1.Enabled = False
        WText2.Enabled = False
         
        CHXFrame.Enabled = False
        CHXLabel.Enabled = False
        CHXText.Enabled = False
    ElseIf temp = 3 Then
        CFrame.Enabled = False
        CText1.Enabled = False
        CText2.Enabled = False
        CLabel1.Enabled = False
        CLabel2.Enabled = False
    
        TSFrame.Enabled = False
        TSLabel1.Enabled = False
        TSLabel2.Enabled = False
        TSText1.Enabled = False
        TSText2.Enabled = False
    
        HFrame.Enabled = True
        HLabel1.Enabled = True
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HLabel7.Enabled = True
        HLabel8.Enabled = True
        HLabel9.Enabled = True
        HText1.Enabled = True
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        HText6.Enabled = True
        HText7.Enabled = True
        HText8.Enabled = True
        HCheck1.Enabled = True
        HCheck2.Enabled = True
        HCheck1.Value = 1
                    
        FHXFrame.Enabled = False
        FHXLabel.Enabled = False
        FHXText.Enabled = False
    
        WFrame.Enabled = False
        WLabel1.Enabled = False
        WLabel2.Enabled = False
        WText1.Enabled = False
        WText2.Enabled = False
         
        CHXFrame.Enabled = False
        CHXLabel.Enabled = False
        CHXText.Enabled = False
    ElseIf temp = 4 Then
        CFrame.Enabled = False
        CText1.Enabled = False
        CText2.Enabled = False
        CLabel1.Enabled = False
        CLabel2.Enabled = False
    
        TSFrame.Enabled = False
        TSLabel1.Enabled = False
        TSLabel2.Enabled = False
        TSText1.Enabled = False
        TSText2.Enabled = False
    
        HFrame.Enabled = False
        HLabel1.Enabled = False
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HLabel7.Enabled = False
        HLabel8.Enabled = False
        HLabel9.Enabled = False
        HText1.Enabled = False
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        HText6.Enabled = False
        HText7.Enabled = False
        HText8.Enabled = False
        HCheck1.Enabled = False
        
        FHXFrame.Enabled = False
        FHXLabel.Enabled = False
        FHXText.Enabled = False
    
        WFrame.Enabled = False
        WLabel1.Enabled = False
        WLabel2.Enabled = False
        WText1.Enabled = False
        WText2.Enabled = False
         
        CHXFrame.Enabled = False
        CHXLabel.Enabled = False
        CHXText.Enabled = False
    Else
        CFrame.Enabled = False
        CText1.Enabled = False
        CText2.Enabled = False
        CLabel1.Enabled = False
        CLabel2.Enabled = False
    
        TSFrame.Enabled = False
        TSLabel1.Enabled = False
        TSLabel2.Enabled = False
        TSText1.Enabled = False
        TSText2.Enabled = False
    
        HFrame.Enabled = False
        HLabel1.Enabled = False
        HLabel2.Enabled = False
        HLabel3.Enabled = False
        HLabel4.Enabled = False
        HLabel5.Enabled = False
        HLabel6.Enabled = False
        HLabel7.Enabled = False
        HLabel8.Enabled = False
        HLabel9.Enabled = False
        HText1.Enabled = False
        HText2.Enabled = False
        HText3.Enabled = False
        HText4.Enabled = False
        HText5.Enabled = False
        HText6.Enabled = False
        HText7.Enabled = False
        HText8.Enabled = False
        HCheck1.Enabled = False
        HCheck2.Enabled = False
    
        FHXFrame.Enabled = False
        FHXLabel.Enabled = False
        FHXText.Enabled = False
    
        WFrame.Enabled = False
        WLabel1.Enabled = False
        WLabel2.Enabled = False
        WText1.Enabled = False
        WText2.Enabled = False
         
        CHXFrame.Enabled = False
        CHXLabel.Enabled = False
        CHXText.Enabled = False
    End If
End Sub




Private Sub SystemOk_Click()
    SimSelect = ModelCombo.ListIndex
    
    If SimSelect = "" Then
        MsgBox "Please select a Simulation Type"
    Else
        If HCheck1.Value = True Then
            UAh = Val(HText1.Text)
        Else
            UAh = Val(HText2.Text) + Val(HText3.Text) + Val(HText4.Text) + Val(HText5.Text)
        End If
        
        Ts = Val(TSText1.Text)
        ms = Val(TSText2.Text) * 0.9
        
        Tc1 = Val(CText1.Text)
        mc = Val(CText2.Text)
        
        e_hx = Val(CHXText.Text)
        
        Qelec = Val(HText6.Text)
        
        Tr = Val(HText7.Text)
        TsupH = Val(HText8.Text)
        
        e_loadhx = Val(FHXText.Text)
        
        Tg = Val(WText1.Text)
        TsupW = Val(WText2.Text)
        
        REMain.Enabled = True
        REMain.Simulate.Enabled = True
        System.Hide
    End If
End Sub


