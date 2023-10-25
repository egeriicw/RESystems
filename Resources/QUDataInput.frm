VERSION 5.00
Begin VB.Form QUDataInput 
   BackColor       =   &H8000000B&
   Caption         =   "Solar Collector Data Input"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Placement"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lGamma 
      Caption         =   "Latitude"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "QUDataInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

