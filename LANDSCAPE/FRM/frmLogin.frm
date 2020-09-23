VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1560
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4005
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":08CA
   ScaleHeight     =   921.701
   ScaleMode       =   0  'User
   ScaleWidth      =   3760.477
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "$"
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.ComboBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmLogin.frx":1194
      Left            =   1320
      List            =   "frmLogin.frx":1196
      TabIndex        =   4
      Tag             =   "temp"
      Text            =   "Select User"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Top             =   1140
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2340
      TabIndex        =   3
      Tag             =   "7-30-06"
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label SRN 
      Caption         =   "Label2"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3120
      Picture         =   "frmLogin.frx":1198
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   -1320
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public LoginSucceeded As Boolean
Dim RS As Recordset
Dim FDIR As String



Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'check for correct password
If txtUserName.Text = txtUserName.Tag And txtPassword.Text = txtPassword.Tag Then
'\\\\\\\If Me.tag = "304681960" Or Me.tag = "1126043640" Or Me.tag = "0" Then
If Me.Tag <> SRN.Tag Then
MsgBox "YOU MUST CONTACT TO PYRO SOFT " & vbCrLf & "( MOHAMMAD ASHFAQ SAQI 92-0301-7619520 )", vbOKOnly, "ERROR IN SYSTEM "
Exit Sub
End If

If toNumber(Image1.Tag) < 949 Then
'MsgBox CDate("07 - 30 - 2006")
       ' Me.MousePointer = 99
        '"304681960"
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
         Label1.Visible = True
        LoginSucceeded = True
        
         
         If SRN.Caption = "Y" Then
        ' MAINPSI.sBAR1.Panels(4).Text = "Admin"
        ' MAINPSI.suty.Visible = True
        ' MAINPSI.Report.Visible = True
         Else
        ' MAINPSI.sBAR1.Panels(4).Text = "User"
         End If
        
        If RS.State = adStateOpen Then
        RS.Close
        End If

        RS.Open " Select * from tbl_users", CNADO, adOpenDynamic, adLockPessimistic
        RS.Fields("DateModified") = Date
        RS.Update
        RS.Close
'Load PInvoiceD
       ' Unload Me
       Me.Hide
       DBR.Show vbModal
       
Else
MsgBox "CONTACT WITH PYRO SOFT, TIME PERIOD ENDED. ", vbOKOnly, "(    Muhammad Ashfaq 92-0301-7619520    )"
Exit Sub
End If
    
    Else
        MsgBox "Invalid Password, try again!", , "Login"
 '       Ag.Characters(Ag.Tag).Speak "ACCESS NOT ALLOWED"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
Function GetSerialNumber(strDrive As String) As Long
'
'
  Dim SerialNum As Long
  Dim Res As Long
  Dim Temp1 As String
  Dim Temp2 As String
 '
  Temp1 = String$(255, Chr$(0))
  Temp2 = String$(255, Chr$(0))
  Res = GetVolumeInformation(strDrive, Temp1, _
  Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
 '
  GetSerialNumber = SerialNum
 
End Function

Private Sub Command1_Click()
If RS.State = adStateOpen Then
RS.Close
End If

RS.Open "Select * from tbl_GENERATOR", CNADO, adOpenDynamic, adLockPessimistic
Image1.Tag = RS.Fields("NEXTNO")
RS.Fields("SRNO") = Command1.Caption
RS.Fields("NEXTNO") = toNumber(Image1.Tag) + 1
'Ag.Tag = RS.Fields("CHR")
RS.Update
RS.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case vbKeyF10
Unload Me

End Select

End Sub

Private Sub Form_Load()
Dim objs
Dim Obj
Dim WMI
MakeADOC


Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_BaseBoard")

Set RS = New Recordset

RS.Open " Select * from tbl_users", CNADO, adOpenDynamic, adLockPessimistic
While Not RS.EOF = True
txtUserName.AddItem RS.Fields(1)
RS.MoveNext
Wend
'cmdCancel.tag = rs.Fields("DateModified")
RS.Close

'
If RS.State = adStateOpen Then
RS.Close
End If

RS.Open "Select * from tbl_GENERATOR", CNADO, adOpenDynamic, adLockPessimistic
Image1.Tag = RS.Fields("NEXTNO")
Me.Tag = RS.Fields("SRNO")
RS.Fields("NEXTNO") = toNumber(Image1.Tag) + 1
'Ag.Tag = RS.Fields("CHR")
RS.Update
RS.Close

'///////

'///////

If RS.State = adStateOpen Then
RS.Close
End If

'RS.Open "Select * from tbl_SalesMaster", CNADO, adOpenDynamic, adLockPessimistic
'cmdCancel.Tag = toNumber(RS.RecordCount)
'RS.Close

'txtUserName.AddItem "PSAdmin"
' Geting HD serial
Label1.Visible = False
Me.MousePointer = 1
SRN.Tag = GetSerialNumber("C:\")
Command1.Caption = SRN.Tag
        
' Getting Motherboad Serial
For Each Obj In objs
  SRN.Tag = Obj.SerialNumber
Next
Command1.Caption = SRN.Tag

'


'MsgBox Me.tag
End Sub

Private Sub Form_Unload(CANCEL As Integer)
LoginSucceeded = False
  Set RS = Nothing
  
Set frmLogin = Nothing

End Sub

Private Sub txtUserName_Click()
On Error GoTo err

If RS.State = adStateOpen Then
RS.Close
End If

RS.Open "select * from tbl_users where userid like '" & txtUserName.Text & "'", CNADO, adOpenDynamic, adLockPessimistic
If RS.EOF = True Then
RS.Close
Else
txtUserName.Tag = RS.Fields(1)
SRN.Caption = RS.Fields("Admin")
txtPassword.Tag = Enc.DecryptString(RS.Fields(2))
'MsgBox txtUserName.Tag & "  " & txtPassword.Tag
If IsNull(RS.Fields(4)) Then
Else
FDIR = App.Path & RS.Fields("PICTURE")
Image1.Picture = LoadPicture(FDIR)
End If

'MAINPSI.sBAR1.Tag = RS.Fields(0)

RS.Close
End If

Exit Sub
err:
MsgBox err.Description

End Sub
