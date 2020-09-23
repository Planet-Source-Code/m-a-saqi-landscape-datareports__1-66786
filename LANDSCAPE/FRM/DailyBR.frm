VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form DBR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Report Date Between (micro_ait@yahoo.com)"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   4035
   ClientWidth     =   5925
   Icon            =   "DailyBR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPort 
      Caption         =   "Portrait"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton optLand 
      Caption         =   "Landscape"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangePrinterOrient 
      Caption         =   "Change Printer Orientation"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1800
      Top             =   2160
   End
   Begin MSComCtl2.DTPicker DateT 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22740992
      CurrentDate     =   38920
   End
   Begin MSComCtl2.DTPicker DateF 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22740992
      CurrentDate     =   37622
      MaxDate         =   2957735
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label SEL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DAILY BUSINESS REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Tag             =   "PURCHASES"
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -600
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu login 
         Caption         =   "Login"
      End
      Begin VB.Menu Worksheet 
         Caption         =   "WorkSheet"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu abtus 
         Caption         =   "About Us"
      End
      Begin VB.Menu shortcuts 
         Caption         =   "&Shotcuts"
      End
   End
End
Attribute VB_Name = "DBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


#If Win32 Then
   Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#ElseIf Win16 Then
   Private Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
   Private Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
#End If
'
' Create searchable listbox objects
'
'Private cLstWin As CListSearch
'Private cLstSys As CListSearch

Dim RS As Recordset
Dim sqls As String
Dim i As Integer

Private Sub abtus_Click()
ABOUTUS.Show vbModal

End Sub

Private Sub CancelButton_Click()
Set RS = Nothing

Unload Me
End
End Sub

Private Sub DGTEMP_Click()
'LedgerSet

End Sub

Private Sub cmdChangePrinterOrient_Click()
   If optPort.Value = True Then
        ChngPrinterOrientationPortrait Me
    ElseIf optLand.Value = True Then
        ChngPrinterOrientationLandscape Me
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case vbKeyF10
Unload Me

End Select

End Sub

Private Sub Form_Load()


Set RS = New Recordset




'Set cCboWin = New CComboSearch
'Set cCboWin.Client = SName
 ' DateF.Value = Date
   DateT.Value = Date + 1
  ' SName.Text = "All Parties"
'SETParties


End Sub

Private Sub Form_Unload(CANCEL As Integer)
If Timer1.Tag = "N" Then
optPort.Value = True
cmdChangePrinterOrient_Click
End If

Set RS = Nothing
Set DBR = Nothing
End
End Sub

Private Sub SETParties()
Set RS = New Recordset

'
If RS.State = adStateOpen Then
RS.Close
End If

'
DateT.Value = Date + 1

'SQLS = "SELECT * From DAILYBR WHERE date between #" & CDate(DateF.Value) & "# and #" & CDate(DateT.Value) & "#"
sqls = "SELECT * From DAILYBr WHERE date >= #" & CDate(DateF.Value) & "# and DATE <= #" & CDate(DateT.Value) & "#"


RS.Open sqls, CNADO, adOpenDynamic, adLockReadOnly
    
    With DBRPT
        .Sections(2).Controls.Item("LBLFROM").Caption = Format$(DateF.Value, "MMM-dd-yyyy")
        .Sections(2).Controls.Item("LBLTO").Caption = Format$(DateT.Value, "MMM-dd-yyyy")
        'rptPurchases.Sections(2).Controls.Item("LBLPADDRESS").Caption = ".............----------............." 'RS.Fields("Address")
        'rptPurchases.Sections(2).Controls.Item("LBLPADDRESS").Caption = RS.Fields("Address")
        Set .DataSource = RS
      
        End With

'RS.Close

End Sub

Private Sub login_Click()
Me.Hide
frmLogin.Show vbModal
End Sub

Private Sub OKButton_Click()
'On Error Resume Next

Call SETPartiesT

Call SETParties

frmChangePrinterOrient.Show vbModal
End Sub





Private Sub shortcuts_Click()
frmshourtkeys.Show vbModal

End Sub

Private Sub Timer1_Timer()

If Timer1.Tag = "N" Then
optLand.Value = True
Else
optPort.Value = True
End If

cmdChangePrinterOrient_Click
'OKButton_Click
Timer1.Enabled = False

End Sub

Private Sub LedgerSet()


End Sub
Private Sub SETPartiesT()
On Error GoTo ERRPT

Set RS = New Recordset

optLand.Value = True
cmdChangePrinterOrient_Click

'
   
If RS.State = adStateOpen Then
RS.Close
End If

'
DateT.Value = Date + 1

'SQLS = "SELECT * From DAILYBR WHERE date between #" & CDate(DateF.Value) & "# and #" & CDate(DateT.Value) & "#"
sqls = "SELECT * From TBL_TRANSACTION WHERE ACCODE not LIKE ''  AND date >= #" & CDate(DateF.Value) & "# and DATE <= #" & CDate(DateT.Value) & "# ORDER BY NAME"


RS.Open sqls, CNADO, adOpenDynamic, adLockReadOnly
    
    
    With DBRPTNew
        .Sections(2).Controls.Item("LBLFROM").Caption = Format$(DateF.Value, "MMM-dd-yyyy")
        .Sections(2).Controls.Item("LBLTO").Caption = Format$(DateT.Value, "MMM-dd-yyyy")
        'rptPurchases.Sections(2).Controls.Item("LBLPADDRESS").Caption = ".............----------............." 'RS.Fields("Address")
        'rptPurchases.Sections(2).Controls.Item("LBLPADDRESS").Caption = RS.Fields("Address")
        Set .DataSource = RS
         End With

'RS.Close

Exit Sub
ERRPT:
MsgBox err.Description

End Sub

Private Sub Worksheet_Click()
If RS.State = adStateOpen Then
RS.Close
End If

'
sqls = "SELECT * From TBL_WORKSHEET"


RS.Open sqls, CNADO, adOpenDynamic, adLockReadOnly

Set WorkSheetRpt.DataSource = RS
WorkSheetRpt.Show vbModal
RS.Close

  
  
End Sub
