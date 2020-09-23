VERSION 5.00
Begin VB.Form frmChangePrinterOrient 
   Caption         =   "Select Report "
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3675
   Icon            =   "ChangePrinterOrient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangePrinterOrient 
      Caption         =   "View"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton optLand 
      Caption         =   "Parties Daily Business Report"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Accounts Daily Business Report"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmChangePrinterOrient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdChangePrinterOrient_Click()
      
    If optPort.Value = True Then
        ChngPrinterOrientationPortrait Me
        DBRPT.Show vbModal
    ElseIf optLand.Value = True Then
        ChngPrinterOrientationLandscape Me
        DBRPTNew.Show vbModal
    End If
        
End Sub

