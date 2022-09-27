Attribute VB_Name = "DextopModule"
Public db As Database

Public Sub initialisePublicVariables()
    Set db = OpenDatabase("Storage.mdb", False, False, "MS Access;PWD=12345abcde")
End Sub


Public Function getNewAccountCode() As String
Dim rs As Recordset, sAccountCode As String
    
    Set rs = db.OpenRecordset("Select Max(val(AccountMaster.Code))As ACode From AccountMaster")
    If rs.RecordCount > 0 Then
        sAccountCode = Val("" & rs!ACode) + 1
    Else
        sAccountCode = "1"
    
    End If
    rs.Close
    
    getNewAccountCode = sAccountCode
End Function

Public Function getNewTransactionForAccount() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( AccountRegister.TransactionNo)) As TNo From AccountRegister ")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionForAccount = sTransactionNo
End Function

Public Function getCurrentBalanceOf(sAccountCode As String) As Double
Dim rs As Recordset
Dim dCurrentBalance As Double
    Set rs = db.OpenRecordset("Select (Sum(AccountRegister.Income)- Sum(AccountRegister.Expense)) As Balance From AccountRegister Where (AccountRegister.AccountCode = '" & sAccountCode & "')")
    If rs.RecordCount > 0 Then
        dCurrentBalance = Val("" & rs!Balance)
    Else
        dCurrentBalance = 0
    End If
    getCurrentBalanceOf = dCurrentBalance
End Function
