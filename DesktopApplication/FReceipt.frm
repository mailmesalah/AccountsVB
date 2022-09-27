VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FReceipt 
   Caption         =   "Accounts - Receipt"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CDelete 
      Height          =   435
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1860
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3990
      Width           =   1485
   End
   Begin VB.CommandButton CSave 
      Height          =   435
      Left            =   3195
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3990
      Width           =   1485
   End
   Begin VB.CommandButton CPrint 
      Height          =   435
      Left            =   1665
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3990
      Width           =   1485
   End
   Begin VB.CommandButton CNew 
      Height          =   435
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3990
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   330
      Left            =   2865
      TabIndex        =   1
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   64028675
      CurrentDate     =   40458
   End
   Begin MSForms.Label LCurrentBalance 
      Height          =   375
      Left            =   1395
      TabIndex        =   13
      Top             =   2655
      Width           =   4770
      VariousPropertyBits=   8388627
      Caption         =   "Current Balance"
      Size            =   "8414;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   375
      Left            =   1410
      TabIndex        =   12
      Top             =   1680
      Width           =   1785
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "3149;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   330
      Left            =   2865
      TabIndex        =   4
      Top             =   1680
      Width           =   3240
      VariousPropertyBits=   746604571
      Size            =   "5715;582"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   375
      Left            =   1410
      TabIndex        =   11
      Top             =   1260
      Width           =   1785
      VariousPropertyBits=   8388627
      Caption         =   "Amount"
      Size            =   "3149;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAmount 
      Height          =   330
      Left            =   2865
      TabIndex        =   3
      Top             =   1260
      Width           =   1710
      VariousPropertyBits=   746604571
      Size            =   "3016;582"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   645
      TabIndex        =   10
      Top             =   180
      Width           =   465
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TVoucherNo 
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   180
      Width           =   1395
      VariousPropertyBits=   746604571
      Size            =   "2461;582"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   375
      Left            =   1410
      TabIndex        =   9
      Top             =   825
      Width           =   1785
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "3149;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   330
      Left            =   2865
      TabIndex        =   2
      Top             =   840
      Width           =   3240
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5715;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sAccountCode() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset

    If Trim(TVoucherNo.Text) = "" Then
        MsgBox "Please Enter a Transaction No !", vbInformation
        TVoucherNo.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionNo = '" & Trim(TVoucherNo.Text) & "' ) And (AccountRegister.Type = 'R' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "Transaction No Not Available !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted !", vbInformation
    clearControls
End Sub

Private Sub CNew_Click()
    clearControls
End Sub

Private Sub CoAccount_Change()
    LCurrentBalance.Caption = "Current Balance is " & Format(getCurrentBalanceOf(sAccountCode(CoAccount.ListIndex + 1)), "0.00")
End Sub

Private Sub CoAccount_GotFocus()
    CoAccount.SelStart = 0
    CoAccount.SelLength = Len(CoAccount.Text)
End Sub

Private Sub CPrint_Click()
    printReceipt
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim sStatus As String

    If Trim(TVoucherNo.Text) = "" Then
        MsgBox "Please give a Transaction No to Edit or Click New to Add new !", vbInformation
        CNew.SetFocus
        Exit Sub
    End If
    
    If CoAccount.ListIndex = -1 Then
        MsgBox "Please Select an Account !", vbInformation
        CoAccount.SetFocus
        Exit Sub
    End If
    
    If Val("" & TAmount.Text) <= 0 Then
        MsgBox "Please Enter valid Amount !", vbInformation
        TAmount.SetFocus
        Exit Sub
    End If
    
    If Trim(TNarration.Text) = "" Then
        MsgBox "Please Enter valid Narration !", vbInformation
        TNarration.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionNo = '" & Trim(TVoucherNo.Text) & "' ) And (AccountRegister.Type = 'R' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
    Else
        sStatus = "Added"
        TVoucherNo.Text = getNewTransactionNo()
        rs.AddNew
        rs!TransactionNo = "" & TVoucherNo.Text
        rs!Type = "R"
    End If

    rs!AccountCode = sAccountCode(CoAccount.ListIndex + 1)
    rs!TransactionDate = DTPDate.Value
    rs!TransactionTime = Format(Time, "HH:MM AMPM")
    rs!Income = Val(TAmount.Text)
    rs!Expense = 0
    rs!Narration = "" & TNarration.Text
    rs.Update
    rs.Close
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    clearControls
    TVoucherNo.SetFocus
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    End If
End Sub

Private Sub Form_Load()
    
    TVoucherNo = getNewTransactionNo
    DTPDate.Value = Date
    getAccountsToCombo
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String

    Set rs = db.OpenRecordset("Select Max(Val(AccountRegister.TransactionNo)) As TNo From AccountRegister Where (AccountRegister.Type = 'R' )")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getAccountsToCombo()
Dim rs As Recordset
    
    CoAccount.Clear
    
    Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.Type = 'BAccount' And AccountMaster.Status = True ) Order By AccountMaster.AccountName")
    ReDim sAccountCode(rs.RecordCount) As String
    While rs.EOF = False
        CoAccount.AddItem "" & rs!AccountName
        sAccountCode(CoAccount.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getTransactionDetails(sTransactionNo As String)
Dim rs As Recordset, sParentCode As String, sAccountCode As String
Dim r As Long

    Set rs = db.OpenRecordset("Select AccountRegister.*,AccountMaster.AccountName From AccountRegister,AccountMaster Where (AccountRegister.TransactionNo = '" & Trim(sTransactionNo) & "' ) And (AccountRegister.Type = 'R' ) And (AccountMaster.Code=AccountRegister.AccountCode)")
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!TransactionDate
        TAmount.Text = Val("" & rs!Income)
        TNarration.Text = "" & rs!Narration
        CoAccount.Text = "" & rs!AccountName
    Else
    
    End If
    rs.Close
End Sub

Private Sub clearControls()
    TVoucherNo.Text = getNewTransactionNo
    DTPDate.Value = Date
    CoAccount.Text = ""
    TAmount.Text = ""
    TNarration.Text = ""
    LCurrentBalance.Caption = ""
End Sub

Private Sub TAmount_GotFocus()
    TAmount.SelStart = 0
    TAmount.SelLength = Len(TAmount.Text)
End Sub

Private Sub TNarration_GotFocus()
    TNarration.SelStart = 0
    TNarration.SelLength = Len(TNarration.Text)
End Sub

Private Sub TVoucherNo_GotFocus()
    TVoucherNo.SelStart = 0
    TVoucherNo.SelLength = Len(TVoucherNo.Text)
End Sub

Private Sub TVoucherNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        getTransactionDetails (TVoucherNo.Text)
    End If
End Sub

Private Sub printReceipt()

    'Dim i, j, x, y As Double
   
    'Printer.ScaleMode = 1
    'Printer.FontName = "Arial"
    
    'Printer.FontBold = True
    'Printer.FontSize = 20
    'Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("DYNAMIC DIGITAL SPOT")) / 2)
    'Printer.CurrentY = 400
    'Printer.Print "DYNAMIC DIGITAL SPOT"
    
    'x = 400
    'y = 900
    
    'Printer.FontSize = 12
    'Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Receipt")) / 2)
    'Printer.CurrentY = y
    'Printer.Print "Receipt"
    
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'y = y + 500
    'Printer.CurrentY = y
    'Printer.Print "No"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(TVoucherNo.Text)
    
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.FontUnderline = False
    'Printer.CurrentX = 4000
    'Printer.CurrentY = y
    'Printer.Print Format(DTPDate.Value, "dd-MMM-yyyy")
    
    'x = 400
    'y = y + 400
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.FontUnderline = False
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Account"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(CoAccount.Text)
    
    'x = 400
    'y = y + 400
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Amount"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Format(TAmount.Text, "0.00")
    
    'x = 400
    'y = y + 400
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Narration"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(TNarration.Text & " " & CoSubAccount.Text)
    
    'Printer.EndDoc
    
    'MsgBox "Successfully send to Printer !", vbInformation
    
    DoEvents    'will not wait to complete the printing,lets to do other things while printing
    
    Dim i As Long, lines As Long, lReturnValue As Long
    
    'checking if the data is already entered
    On Error GoTo GoOut
    
    Open "LPT1:" For Output As #1
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "    " & Chr(0) & Chr(27) & "!" & Chr(50) & "DeXtop" & Chr(27) & "!" & Chr(0) & Chr(27) & "!" & Chr(20) & " Software Innovations" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "    " & "Receipt" & Chr(0)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "No        : R/" & Left(Trim(TVoucherNo.Text & "") & Space(22), 22) & Space(90) & " Date: " & Left(Format(DTPDate.Value, "dd-MMM-yyyy") & Space(12), 12) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "Account   : " & Left(CoAccount.Text & Space(40), 40)
    Print #1, Chr(27) & "!" & Chr(20) & "Narration : " & Left(TNarration.Text & Space(40), 40)
    Print #1, Chr(27) & "!" & Chr(20) & "Amount    : " & Right(Space(12) & Format(TAmount.Text, "0.00"), 12)
    'Print #1, Chr(27) & "!" & Chr(20) & "Narration : " & Left(Trim(TNarration.Text & " " & CoSubAccount.Text) & Space(40), 40)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""

    Close #1
    
    lReturnValue = MsgBox("Successfully Send to Printed !", vbInformation)
    
    Exit Sub
GoOut:
    MsgBox "Check If Printer is available, " & Err.Description, vbInformation
End Sub
