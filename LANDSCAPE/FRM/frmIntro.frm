VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      Picture         =   "frmIntro.frx":20082
      ScaleHeight     =   5325.001
      ScaleMode       =   0  'User
      ScaleWidth      =   5338.327
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -240
         Top             =   3360
      End
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************
'Copyright © 2001 by Alexander Anikin
'e-mail: aka@i.com.ua
'http://hotmix.narod.ru
'*************************************
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = (-12)
Private hOldCursor As Long
Private hNewCursor As Long
     Dim Proba, Proba2 As Integer
     Dim Boja2 As String

Private Function Zrak(slika As PictureBox, StartX As Integer, StartY As Integer, Levo As Integer, Desno As Integer, Boja As String)

     

Me.ScaleMode = vbPixels

With slika
     .ScaleMode = vbPixels
     .AutoRedraw = True
End With

For Proba2 = 0 To slika.ScaleWidth
    DoEvents

For Proba = 0 To slika.ScaleHeight
    Boja2 = slika.Point(Proba2, Proba)
   Line (StartX, StartY)-(Levo + Proba2, Desno + Proba), Boja2
Next
   Line (StartX, StartY)-(Levo + Proba2, Desno + slika.ScaleHeight), Boja
Next

For Proba2 = 0 To slika.ScaleHeight
   Line (StartX, StartY)-(Levo + slika.ScaleWidth, Desno + Proba2), Boja
Next

End Function



Private Sub Form_Load()
hNewCursor = LoadCursorFromFile(App.Path & "\Dsystem\cur.cps")
hOldCursor = SetClassLong(hwnd, GCL_HCURSOR, hNewCursor)

End Sub

Private Sub Form_Unload(CANCEL As Integer)
hOldCursor = SetClassLong(hwnd, GCL_HCURSOR, hOldCursor)
frmLogin.Show
End Sub

Private Sub Timer1_Timer()
Zrak Picture1, 565, 301, 0, 0, Me.BackColor    ' adjust the scale and position from where
                                               ' to the size   of the picture , or perform some variation.
   If Timer1.Interval = 1000 Then
'    frmAdministrator.Show    , if you are going further with your program
                     ' you can replace timer with command_click()
    
    Unload Me
End If
End Sub
                                             
                                                

