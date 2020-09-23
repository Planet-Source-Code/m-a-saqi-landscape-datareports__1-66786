Attribute VB_Name = "PSFIELDS"

'Dim WithEvents adoPrimaryRS As Recordset
Global a As Integer

Global CNADO As Connection
'Global RSADO As New ADODB.Recordset

Global cnSQL As New ADODB.Connection
Global rsSQL As New ADODB.Recordset
Global Enc As New clsBlowfish

Public Function MakeSQLC()
   On Error GoTo ErrHandler:
  Dim UserName As String
  Dim Pass As String
  Dim DBname As String
  Dim ServerN As String
  
  UserName = ""
  Pass = ""
  'DBname = "Northwind"
  DBname = "MADIMAGIC"
  ServerN = "Pyro"
    
  cnSQL.CursorLocation = adUseServer
  
  cnSQL.Provider = "sqloledb"
  cnSQL.Properties("Data Source").Value = ServerN
  cnSQL.Properties("Initial Catalog").Value = DBname
  cnSQL.Properties("Integrated Security").Value = "SSPI"
  cnSQL.Open
  
  'PHT = "C:\psinveNtory\DSYSTEM\PSIS.mdb"
  
  'db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & PHT & ";"
  'Set rs = New Recordset
       
       Exit Function
ErrHandler:
    MsgBox err.Description, , "Error "
End Function
Public Function MakeADOC()
   On Error GoTo ErrHandler:
  Dim Pht As String
  Set CNADO = New Connection
'  Dim Pass As String
'  Dim DBname As String
'  Dim ServerN As String
  
 ' UserName = ""
 ' Pass = ""
 ' DBname = "Northwind"
 ' ServerN = "Pyro"
    
  CNADO.CursorLocation = adUseClient
  
  'cn.Provider = "sqloledb"
  'cn.Properties("Data Source").Value = ServerN
  'cn.Properties("Initial Catalog").Value = DBname
  'cn.Properties("Integrated Security").Value = "SSPI"
  'cn.Open
  
  Pht = App.Path & "\DSYSTEM\db.mdb"
  
  CNADO.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pht & ";Persist Security Info=false; Jet OLEDB:Database Password=ADMIN "
  'Set rs = New Recordset
       
       Exit Function
ErrHandler:
    MsgBox err.Description, , "Error "
End Function

Public Function OpenSQLR()
   On Error GoTo ErrHandler:
  Set rsSQL = New ADODB.Recordset
  
  'rs.Open "SELECT * FROM MPO", db, adOpenStatic, adLockOptimistic
  
     Exit Function
ErrHandler:
    MsgBox err.Description, , "Error "
End Function

Public Function OpenADOR()
   On Error GoTo ErrHandler:
  Dim rsado As ADODB.Recordset
  Set rsado = New ADODB.Recordset
  
  'rs.Open "SELECT * FROM MPO", db, adOpenStatic, adLockOptimistic
  
     Exit Function
ErrHandler:
    MsgBox err.Description, , "Error "
End Function

Public Function CloseC()
 '[  On Error GoTo ErrHandler:
   
   If rsSQL.State = adStateOpen Then
   rsSQL.Close
   End If
   
'   If rsado.State = adStateOpen Then
'   rsado.Close
'   End If
   
   
   If cnSQL.State = adStateOpen Then
  cnSQL.Close
  End If
  
  If CNADO.State = adStateOpen Then
  CNADO.Close
  End If
  '
  
  'rs.Open "SELECT * FROM MPO", db, adOpenStatic, adLockOptimistic
  
   '  Exit Function
'ErrHandler:
 '   MsgBox Err.Description, , "Error "
End Function


'Convert string to number
'I create this istead of val() co'z val return incorrect value
'ex. Try to see the output of val("3,800")
'It did not support characters like , and etc.
Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function
'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0.00")
End Function

