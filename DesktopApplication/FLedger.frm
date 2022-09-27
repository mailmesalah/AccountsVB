VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FLedger 
   Caption         =   "Accounts Report"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   ControlBox      =   0   'False
   Icon            =   "FLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   10185
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7095
      Width           =   2175
   End
   Begin VB.CommandButton CToExcel 
      Caption         =   "To Excel"
      Height          =   570
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7095
      Width           =   2175
   End
   Begin VB.CommandButton CShow 
      Caption         =   "Show"
      Height          =   570
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7095
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   7329
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   8421504
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   1845
      TabIndex        =   18
      Top             =   210
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20578307
      CurrentDate     =   40458
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1845
      TabIndex        =   19
      Top             =   645
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20578307
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   285
      TabIndex        =   21
      Top             =   240
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   285
      TabIndex        =   20
      Top             =   630
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   330
      Left            =   9450
      TabIndex        =   17
      Top             =   690
      Width           =   2790
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4921;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Left            =   7125
      TabIndex        =   16
      Top             =   675
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LBalance 
      Height          =   405
      Left            =   7980
      TabIndex        =   15
      Top             =   7125
      Width           =   1140
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2011;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   420
      Left            =   7125
      TabIndex        =   14
      Top             =   210
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoType 
      Height          =   330
      Left            =   9450
      TabIndex        =   13
      Top             =   225
      Width           =   2790
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4921;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5130
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label LPayment 
      Height          =   405
      Left            =   10605
      TabIndex        =   11
      Top             =   6150
      Width           =   1140
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2011;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LReceipt 
      Height          =   405
      Left            =   9090
      TabIndex        =   10
      Top             =   6165
      Width           =   1140
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2011;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   9060
      TabIndex        =   9
      Top             =   1455
      Width           =   1140
      VariousPropertyBits=   8388627
      Caption         =   "Receipt"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   1365
      TabIndex        =   8
      Top             =   1470
      Width           =   1605
      VariousPropertyBits=   8388627
      Caption         =   "Voucher No"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   180
      TabIndex        =   7
      Top             =   1455
      Width           =   1410
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "2487;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   6585
      TabIndex        =   6
      Top             =   1455
      Width           =   1095
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "1931;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   2970
      TabIndex        =   5
      Top             =   1470
      Width           =   2610
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "4604;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   10575
      TabIndex        =   4
      Top             =   1455
      Width           =   1140
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim gDate As Single, gVoucherNo As Single, gAccount As Single, gDescription As Single, gReceipt As Single, gPayment As Single
Dim sAccountCode() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gVoucherNo = 1
    gAccount = 2
    gDescription = 3
    gReceipt = 4
    gPayment = 5
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 6
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gVoucherNo) = 1200
    MGrid.ColWidth(gAccount) = 2075
    MGrid.ColWidth(gDescription) = 4000
    MGrid.ColWidth(gReceipt) = 1500
    MGrid.ColWidth(gPayment) = 1500
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CoType_LostFocus()
    getAccounts
End Sub

Private Sub CPrint_Click()
    If MGrid.Rows = 0 Then
        MsgBox "Empty Grid !", vbInformation
        Exit Sub
    End If
    'printReport
End Sub

Private Sub getAccounts()
Dim rs As Recordset
    
    CoAccount.Clear
    
    If (CoType.ListIndex = 0) Then
        Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.Type='BAccount')")
    ElseIf (CoType.ListIndex > 0) Then
        Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.Type='AGroup')")
    Else
        Exit Sub
    End If
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sAccountCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoAccount.AddItem "" & rs!AccountName
        sAccountCode(CoAccount.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
    MGrid.Rows = 0
    If (CoType.ListIndex = 0) Then
        If (CoAccount.ListIndex >= 0) Then
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.TransactionNo,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.AccountCode = '" & sAccountCode(CoAccount.ListIndex + 1) & "' ) And (AccountRegister.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Order By AccountRegister.TransactionDate,Val(AccountRegister.TransactionNo)")
        Else
            Exit Sub
        End If
    ElseIf (CoType.ListIndex = 1) Then
        If (CoAccount.ListIndex >= 0) Then
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.TransactionNo,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.AccountCode In (Select A.Code From AccountMaster As A  Where A.GroupCode='" & sAccountCode(CoAccount.ListIndex + 1) & "') ) And (AccountRegister.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Order By AccountRegister.TransactionDate,Val(AccountRegister.TransactionNo)")
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    While rs.EOF = False
        MGrid.AddItem Format("" & rs!TransactionDate, "dd-MM-yyyy") & vbTab & "" & rs!TransactionNo & vbTab & "" & rs!AccountName & vbTab & "" & rs!Narration & vbTab & Format(Val("" & rs!Income), "0.00") & vbTab & Format(Val("" & rs!Expense), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    
    getTotals
End Sub

Private Sub CToExcel_Click()
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
    ReDim xData(1 To lRowCount + 2, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

    xData(1, 1) = "Date"
    xData(1, 2) = "Bill No"
    xData(1, 3) = "Description"
    xData(1, 4) = "Bill Amount"
    xData(1, 5) = "Advance"
    xData(1, 6) = "Balance"

    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    'xData(i + 1, 4) = LBillAmount.Caption
    'xData(i + 1, 5) = LAdvance.Caption
    'xData(i + 1, 6) = LBalance.Caption
    
    oExcelSheet.Range("A3:F" & lRowCount + 4).Value = xData

    'oExcelSheet.Cells(1, 1).Value = "Laser Sale Bill Wise Summary From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:F" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
    getTotals
End Sub

Private Sub DTPTo_Change()
    MGrid.Rows = 0
    getTotals
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase("Storage.mdb", False, False, "MS Access;PWD=12345abcde")
    MGridInitialise
    DTPFrom.Value = Date
    DTPTo.Value = Date
    CoType.AddItem "Single Account"
    CoType.AddItem "Group Account"
End Sub

Private Sub getTotals()
Dim r As Long
Dim dReceipt As Double, dPayment As Double
    r = 0
    dReceipt = 0
    dPayment = 0
    While r < MGrid.Rows
        dReceipt = dReceipt + Val(MGrid.TextMatrix(r, gReceipt))
        dPayment = dPayment + Val(MGrid.TextMatrix(r, gPayment))
        r = r + 1
    Wend
    LReceipt.Caption = Format("" & dReceipt, "0.00")
    LPayment.Caption = Format("" & dPayment, "0.00")
    LBalance.Caption = Format("" & (Val(dPayment) - Val(dReceipt)), "0.00")
End Sub
