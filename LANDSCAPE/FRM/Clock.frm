VERSION 5.00
Begin VB.Form frmshourtkeys 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Short Keys"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Clock.frx":0E42
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   480
      Top             =   2280
   End
End
Attribute VB_Name = "frmshourtkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ymouse, Xmouse, dy(30), dx(30), DA(30), Mo(30)
Dim Split1, Day1, Year1, Todaysdate, h, m, D, S, Face, Speed, N, scrll
Dim Dsplit, HandHeight, Handwidth, HandX, HandY, Step, currStep
Dim Test, ClockHeight, ClockWidth, ClockFromMouseY, ClockFromMouseX
Dim Fcol, Mcol, Scol, Hcol, Dcol
Private Type FL
  T(30) As Long
  Le(30) As Long
End Type
Dim FL As FL
Private Type HL
  T(30) As Long
  Le(30) As Long
End Type
Dim HL As HL
Private Type sl
  T(30) As Long
  Le(30) As Long
End Type
Dim sl As sl
Private Type ML
  T(30) As Long
  Le(30) As Long
End Type
Dim ML As ML
Private Type DL
  T(30) As Long
  Le(30) As Long
End Type
Dim DL As DL
Const Pi = 3.1415
Private Sub Form_Load()
Dcol = 150   '//date colour.
Fcol = vbBlue  '//face colour.
Scol = 0    '//seconds colour.
Mcol = 0   '//minutes colour.
Hcol = 0   '//hours colour.
ClockHeight = 600
ClockWidth = 600
ClockFromMouseY = 1200
ClockFromMouseX = 600
'//Alter nothing below! Alignments will be lost!
DA(1) = "SUNDAY": DA(2) = "MONDAY": DA(3) = "TUESDAY": DA(4) = "WEDNESDAY"
DA(5) = "THURSDAY": DA(6) = "FRIDAY": DA(7) = "SATURDAY"
Mo(1) = "JANUARY": Mo(2) = "FEBRUARY": Mo(3) = "MARCH"
Mo(4) = "APRIL": Mo(5) = "MAY": Mo(6) = "JUNE": Mo(7) = "JULY"
Mo(8) = "AUGUST": Mo(9) = "SEPTEMBER": Mo(10) = "OCTOBER"
Mo(11) = "NOVEMBER": Mo(12) = "DECEMBER"
Day1 = Day(Now)
Year1 = Year(Now)
If (Year1 < 2000) Then Year1 = Year1 + 1900
Todaysdate = " " + DA(Weekday(Now)) + " " + Str(Day1) + " " + Mo(Month(Now)) + " " + Str(Year1)
D = Todaysdate
h = "..."
m = "...."
S = "....."
Face = "1 2 3 4 5 6 7 8 9 101112  "
frmshourtkeys.Font = "Arial"
frmshourtkeys.FontSize = 8
Speed = 0.6
N = Len(Face) - 2
Ymouse = 0
Xmouse = 0
scrll = 0
Split1 = 360 / N
Dsplit = 360 / Len(D)
HandHeight = ClockHeight / 4.5
Handwidth = ClockWidth / 4.5
HandY = -7
HandX = -2.5
scrll = 0 '2 * ClockHeight
Step = 0.06
currStep = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ymouse = Y + ClockFromMouseY ':event.y+ClockFromMouseY;
Xmouse = X + ClockFromMouseX ':event.x+ClockFromMouseX;
End Sub

Private Sub Timer1_Timer()
dy(0) = dy(0) + (Ymouse - dy(0)) * Speed
dy(0) = Fix(dy(0))
dx(0) = dx(0) + (Xmouse - dx(0)) * Speed
dx(0) = Fix(dx(0))
For i = 1 To Len(D) - 1
dy(i) = dy(i) + (dy(i - 1) - dy(i)) * Speed
dy(i) = Fix(dy(i))
dx(i) = dx(i) + (dx(i - 1) - dx(i)) * Speed
dx(i) = Fix(dx(i))
Next i
secs = Second(Now)
sec = -1.57 + Pi * secs / 30
Mins = Minute(Now)
Min = -1.57 + Pi * Mins / 30
hr = Hour(Now)
hrs = -1.575 + Pi * hr / 6 + Pi * Int(Minute(Now)) / 360
For i = 0 To N - 2
 FL.T(i) = dy(i) + ClockHeight * Sin(-1.0471 + i * Split1 * Pi / 180) + scrll
 FL.Le(i) = dx(i) + ClockWidth * Cos(-1.0471 + i * Split1 * Pi / 180)
Next i
For i = 0 To Len(h) - 1
 HL.T(i) = dy(i) + HandY + (i * HandHeight) * Sin(hrs) + scrll
 HL.Le(i) = dx(i) + HandX + (i * Handwidth) * Cos(hrs)
Next i
For i = 0 To Len(m) - 1
 ML.T(i) = dy(i) + HandY + (i * HandHeight) * Sin(Min) + scrll
 ML.Le(i) = dx(i) + HandX + (i * Handwidth) * Cos(Min)
Next i
For i = 0 To Len(S) - 1
 sl.T(i) = dy(i) + HandY + (i * HandHeight) * Sin(sec) + scrll
 sl.Le(i) = dx(i) + HandX + (i * Handwidth) * Cos(sec)
Next i
For i = 0 To Len(D) - 1
 DL.T(i) = dy(i) + ClockHeight * 1.5 * Sin(currStep + i * Dsplit * Pi / 180) + scrll
 DL.Le(i) = dx(i) + ClockWidth * 1.5 * Cos(currStep + i * Dsplit * Pi / 180)
Next i
currStep = currStep - Step
p
End Sub

Private Function sp(ByVal ST As String, ByVal Nu As Integer, Optional K As Byte = 1) As String
sp = Mid(ST, Nu + 1, K)
End Function
Private Sub p()
Cls
With frmshourtkeys
.FontBold = False
.ForeColor = Dcol
For i = 0 To Len(D) - 1
.CurrentY = DL.T(i)
.CurrentX = DL.Le(i)
Print sp(D, i)
Next i
.ForeColor = Fcol
For i = 0 To N - 1
.CurrentY = FL.T(i)
.CurrentX = FL.Le(i)
If (i = 18 Or i = 20 Or i = 22) Then
Print sp(Face, i, 2)
i = i + 1
Else
Print sp(Face, i, 1)
End If
Next i
.FontBold = True
.ForeColor = Scol
For i = 0 To Len(S) - 1
.CurrentY = sl.T(i)
.CurrentX = sl.Le(i)
Print sp(S, i)
Next i
.ForeColor = Mcol
For i = 0 To Len(m) - 1
.CurrentY = ML.T(i)
.CurrentX = ML.Le(i)
Print sp(m, i)
Next i
.ForeColor = Hcol
For i = 0 To Len(h) - 1
.CurrentY = HL.T(i)
.CurrentX = HL.Le(i)
Print sp(h, i)
Next i
End With
End Sub

