VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FJournalVoucher 
   Caption         =   "Day Transaction"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   285
      Width           =   1545
   End
   Begin VB.CommandButton CClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1545
   End
   Begin VB.CommandButton CRemoveItem 
      Caption         =   "Remove Item"
      Height          =   435
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1545
   End
   Begin VB.CommandButton CAddItem 
      Caption         =   "Add Item"
      Height          =   435
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1545
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   570
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   570
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton CExport 
      Caption         =   "Export To Excel"
      Height          =   570
      Left            =   2355
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8640
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3660
      Left            =   285
      TabIndex        =   1
      Top             =   1875
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   6456
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   2415
      TabIndex        =   0
      Top             =   300
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label3 
      Height          =   375
      Left            =   420
      TabIndex        =   27
      Top             =   345
      Width           =   465
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   960
      TabIndex        =   26
      Top             =   300
      Width           =   1410
      VariousPropertyBits=   746604571
      Size            =   "2487;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   6375
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   4740
      TabIndex        =   24
      Top             =   1545
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   4380
      TabIndex        =   3
      Top             =   5685
      Width           =   3300
      VariousPropertyBits=   746604571
      Size            =   "5821;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Left            =   8085
      TabIndex        =   23
      Top             =   7305
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalance 
      Height          =   345
      Left            =   9570
      TabIndex        =   22
      Top             =   7305
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   420
      Left            =   8040
      TabIndex        =   21
      Top             =   6450
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Total Receipt"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalPayment 
      Height          =   345
      Left            =   9570
      TabIndex        =   20
      Top             =   6855
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Total Payment0"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   8085
      TabIndex        =   19
      Top             =   6855
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Total Payment"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalReceipt 
      Height          =   420
      Left            =   9570
      TabIndex        =   18
      Top             =   6480
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Total Receipt0"
      Size            =   "2725;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   9555
      TabIndex        =   17
      Top             =   1545
      Width           =   1560
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   7980
      TabIndex        =   16
      Top             =   1545
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Receipt"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1755
      TabIndex        =   15
      Top             =   1545
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   330
      TabIndex        =   14
      Top             =   1545
      Width           =   1050
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "1852;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TPayment 
      Height          =   420
      Left            =   9480
      TabIndex        =   5
      Top             =   5685
      Width           =   1710
      VariousPropertyBits=   746604571
      Size            =   "3016;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TReceipt 
      Height          =   420
      Left            =   7725
      TabIndex        =   4
      Top             =   5685
      Width           =   1710
      VariousPropertyBits=   746604571
      Size            =   "3016;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoAccounts 
      Height          =   420
      Left            =   1560
      TabIndex        =   2
      Top             =   5685
      Width           =   2775
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4895;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LSlNo 
      Height          =   390
      Left            =   285
      TabIndex        =   13
      Top             =   5700
      Width           =   1155
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "2037;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FJournalVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sAccountCode() As String
Dim gSlNo As Single, gAccount As Single, gNarration As Single, gReceipt As Single, gPayment As Single, gAccountCode As Single

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If CoAccounts.ListIndex = -1 Then
        MsgBox "Please Select an Account !", vbInformation
        CoAccounts.SetFocus
        Exit Sub
    End If
    
    If Val(TReceipt.Text) = 0 And Val(TPayment.Text) = 0 Then
        MsgBox "Please Enter Receipt or Payment !", vbInformation
        TReceipt.SetFocus
        Exit Sub
    End If
    
    If Val(TReceipt.Text) > 0 And Val(TPayment.Text) > 0 Then
        MsgBox "Please Enter Receipt or Payment only one at a time !", vbInformation
        TReceipt.SetFocus
        Exit Sub
    End If
        
    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
    
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSlNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(r - 1, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(r - 1, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(r - 1, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(r - 1, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
    End If
    clearEditControls
    setBalance
    CoAccounts.SetFocus
End Sub

Private Sub CClear_Click()
    MGrid.Rows = 0
    setBalance
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSlNo = 0
    gAccount = 1
    gNarration = 2
    gReceipt = 3
    gPayment = 4
    gAccountCode = 5
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 6
    MGrid.Rows = 0
    MGrid.ColWidth(gSlNo) = 1200
    MGrid.ColWidth(gAccount) = 2965
    MGrid.ColWidth(gNarration) = 3300
    MGrid.ColWidth(gReceipt) = 1700
    MGrid.ColWidth(gPayment) = 1700
    MGrid.ColWidth(gAccountCode) = 0
    MGrid.RowHeightMin = 350
End Sub

Public Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( AccountRegister.TransactionNo)) As TNo From AccountRegister Where (AccountRegister.Type='JV') ")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getAccounts()
Dim rs As Recordset
    
    CoAccounts.Clear
    
    Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.Status = True ) And (AccountMaster.Type = 'BAccount' )")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sAccountCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoAccounts.AddItem "" & rs!AccountName
        sAccountCode(CoAccounts.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub clearControls()
    LSlNo.Caption = MGrid.Rows + 1
    DTPDate.Value = Date
    MGrid.Rows = 0
    CoAccounts.ListIndex = -1
    TNarration.Text = ""
    TReceipt.Text = ""
    TPayment.Text = ""
    TTransactionNo.Text = getNewTransactionNo
    setBalance
End Sub

Private Sub clearEditControls()
    CoAccounts.ListIndex = -1
    TNarration.Text = ""
    TReceipt.Text = ""
    TPayment.Text = ""
    LSlNo.Caption = MGrid.Rows + 1
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete this day's Transaction ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionNo ='" & Val(TTransactionNo.Text) & "') And (AccountRegister.Type='JV')")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        If bFound Then
            MsgBox "Successfully Deleted !", vbInformation
            clearControls
            getTransactionDetails
            TTransactionNo.Text = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
End Sub

Private Sub CExport_Click()
On Error GoTo ErrHandler
Dim oExcel As Object, oExcelSheet As Object
Dim lReturnValue As Long
Dim lRowCount As Long, lColCount As Long

    If MGrid.Rows = 0 Then
        MsgBox "Empty Data!", vbInformation
        Exit Sub
    End If
  
    OLEExcel.CreateEmbed vbNullString, "Excel.Sheet"
    
    lRowCount = MGrid.Rows
    lColCount = MGrid.Cols
    ReDim xData(1 To lRowCount + 3, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

    xData(1, 1) = "Voucher No"
    xData(1, 2) = "Account"
    xData(1, 3) = "Narration"
    xData(1, 4) = "Receipt"
    xData(1, 5) = "Payment"
    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 4) = Format(LTotalReceipt.Caption, "0.00")
    xData(i + 1, 5) = Format(LTotalPayment.Caption, "0.00")
    xData(i + 2, 4) = "Balance"
    xData(i + 2, 5) = Format(LBalance.Caption, "0.00")
    
    oExcelSheet.Range("A3:E" & lRowCount + 5).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Journal Voucher of " & Format(DTPDate.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:E" & lRowCount + 5).Select
    oExcel.Application.Selection.AutoFormat

On Error Resume Next
    ' Delete the existing test file (if any)...
    Kill App.Path & "\Reports\JournalVoucher " & Format(DTPDate.Value, "dd-MMM-yyyy") & ".xlsx"

  ' Save the file as a native XLS file...
    oExcel.SaveAs App.Path & "\Reports\JournalVoucher " & Format(DTPDate.Value, "dd-MMM-yyyy") & ".xlsx"
    
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
  ' Close the OLE object and remove it...
    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\JournalVoucher " & Format(DTPDate.Value, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

End Sub

Private Sub CoAccounts_GotFocus()
    CoAccounts.SelStart = 0
    CoAccounts.SelLength = Len(CoAccounts.Text)
End Sub

Private Sub CoAccounts_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FAccountMaster.Show vbModal
        getAccounts
    End If
End Sub

Private Sub CRemoveItem_Click()
Dim r As Long
    If MGrid.Rows <= 0 Then
        Exit Sub
    End If
    
    If MGrid.Rows = 1 Then
        MGrid.Rows = 0
        clearEditControls
    Else
        MGrid.RemoveItem (MGrid.Row)
        clearEditControls
    End If
    setBalance
    
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim r As Long, lYN As Long, sStatus As String

    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoAccounts.SetFocus
        Exit Sub
    End If
    
    If Val(LBalance.Caption) <> 0 Then
        MsgBox "The balance has to be zero !", vbInformation
        CoAccounts.SetFocus
        Exit Sub
    End If
    'SAVES DATA TO AccountRegister TABLE
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And (AccountRegister.Type='JV') )")
        
    'SAVES DATA TO TransactionRegister ReadyMade
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    r = 0
    While r < MGrid.Rows
        rs.AddNew
        rs!TransactionNo = Val(TTransactionNo.Text)
        rs!Type = "JV"
        rs!TransactionDate = DTPDate.Value
        rs!TransactionTime = Format(Time, "HH:MM AMPM")
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSlNo))
        rs!AccountCode = Trim(MGrid.TextMatrix(r, gAccountCode))
        rs!Narration = Trim(MGrid.TextMatrix(r, gNarration))
        rs!Income = Val(MGrid.TextMatrix(r, gReceipt))
        rs!Expense = Val(MGrid.TextMatrix(r, gPayment))
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    MsgBox "Successfully Saved !", vbInformation
    clearControls
    getTransactionDetails
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CExport_Click
    ElseIf (KeyCode = vbKeyE And ((Shift And 7) = 2)) Then
        CAddItem_Click
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        CRemoveItem_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClear_Click
    End If
End Sub

Private Sub Form_Load()

    getAccounts
    MGridInitialise
    clearControls
    getTransactionDetails
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long
    
    If MGrid.Rows <= 0 Then
        Exit Sub
    End If
    
    r = MGrid.Row
    LSlNo.Caption = Val(MGrid.TextMatrix(r, gSlNo))
    TNarration.Text = Trim(MGrid.TextMatrix(r, gNarration))
    TReceipt.Text = Val(MGrid.TextMatrix(r, gReceipt))
    TPayment.Text = Val(MGrid.TextMatrix(r, gPayment))
    CoAccounts.Text = Trim(MGrid.TextMatrix(r, gAccount))

End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TReceipt_GotFocus()
    TReceipt.SelStart = 0
    TReceipt.SelLength = Len(TReceipt.Text)
End Sub

Private Sub TPayment_GotFocus()
    TPayment.SelStart = 0
    TPayment.SelLength = Len(TPayment.Text)
End Sub

Private Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select AccountRegister.*,AccountMaster.AccountName From AccountRegister,AccountMaster Where ((AccountRegister.TransactionNo = '" & Val(TTransactionNo.Text) & "') And (AccountMaster.Code=AccountRegister.AccountCode ) And (AccountRegister.Type='JV')) Order By Val(AccountRegister.SerialNo)")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = DateValue("" & rs!TransactionDate)
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSlNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gAccountCode) = "" & rs!AccountCode
            MGrid.TextMatrix(r, gAccount) = "" & rs!AccountName
            MGrid.TextMatrix(r, gNarration) = "" & rs!Narration
            MGrid.TextMatrix(r, gReceipt) = "" & rs!Income
            MGrid.TextMatrix(r, gPayment) = "" & rs!Expense
            
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    LSlNo.Caption = MGrid.Rows + 1
    setBalance
End Sub

Private Sub setBalance()
    getTotalReceiptPayment
    LBalance.Caption = Format(Val(LTotalReceipt.Caption) - Val(LTotalPayment.Caption), "0.00")
End Sub

Private Sub getTotalReceiptPayment()
Dim r As Long, dReceipt As Double, dPayment As Double
    r = 0
    dReceipt = 0
    dPayment = 0
    While r < MGrid.Rows
        dReceipt = dReceipt + Val(MGrid.TextMatrix(r, gReceipt))
        dPayment = dPayment + Val(MGrid.TextMatrix(r, gPayment))
        r = r + 1
    Wend
    LTotalReceipt.Caption = Format(dReceipt, "0.00")
    LTotalPayment.Caption = Format(dPayment, "0.00")
End Sub

Private Sub TTransactionNo_GotFocus()
    TTransactionNo.SelStart = 0
    TTransactionNo.SelLength = Len(TTransactionNo.Text)
End Sub

Private Sub TTransactionNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        getTransactionDetails
    End If
End Sub
