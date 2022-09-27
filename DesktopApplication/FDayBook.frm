VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FDayBook 
   Caption         =   "Day Book"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   ControlBox      =   0   'False
   Icon            =   "FDayBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   10185
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7125
      Width           =   2175
   End
   Begin VB.CommandButton CToExcel 
      Height          =   570
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7125
      Width           =   2175
   End
   Begin VB.CommandButton CShow 
      Height          =   570
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7125
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   1830
      TabIndex        =   0
      Top             =   180
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20709379
      CurrentDate     =   40458
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1830
      TabIndex        =   1
      Top             =   615
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20709379
      CurrentDate     =   40458
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   240
      TabIndex        =   2
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
      GridColorFixed  =   12632256
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
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5235
      TabIndex        =   12
      Top             =   285
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   10665
      TabIndex        =   11
      Top             =   1485
      Width           =   1350
      VariousPropertyBits=   8388627
      Caption         =   "Debit"
      Size            =   "2381;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   9120
      TabIndex        =   10
      Top             =   1485
      Width           =   1350
      VariousPropertyBits=   8388627
      Caption         =   "Credit"
      Size            =   "2381;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   345
      Left            =   1950
      TabIndex        =   9
      Top             =   1485
      Width           =   5490
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "9684;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   315
      TabIndex        =   8
      Top             =   1470
      Width           =   1110
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "1958;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   225
      TabIndex        =   7
      Top             =   180
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
      Left            =   210
      TabIndex        =   6
      Top             =   585
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FDayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim gDate As Single, gToBy As Single, gDescription As Single, gReceipt As Single, gPayment As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gToBy = 1
    gDescription = 2
    gReceipt = 3
    gPayment = 4
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 5
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gToBy) = 900
    MGrid.ColWidth(gDescription) = 6460
    MGrid.ColWidth(gReceipt) = 1500
    MGrid.ColWidth(gPayment) = 1500
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CPrint_Click()
    If MGrid.Rows = 0 Then
        MsgBox "Empty Grid !", vbInformation
        Exit Sub
    End If
    'printReport
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
Dim dDate As Date
Dim dReceipt As Double, dPayment As Double

    MGrid.Rows = 0

    Set rs = db.OpenRecordset("Select AccountRegister.TransactionNo,AccountRegister.AccountCode,AccountRegister.TransactionDate,AccountRegister.Type,AccountRegister.Expense,AccountRegister.Income,AccountRegister.Narration,(Select AccountMaster.AccountName From AccountMaster Where AccountMaster.Code=AccountRegister.AccountCode ) As AccountDescription,(Select (Sum(AccountRegister.Income)-Sum(AccountRegister.Expense)) From AccountRegister Where AccountRegister.TransactionDate<cDate('" & DTPFrom.Value & "')) As OpeningBalance From AccountRegister Where TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') Order By AccountRegister.TransactionDate,AccountRegister.TransactionTime")
    If rs.RecordCount > 0 Then
        MGrid.AddItem ""
        dReceipt = 0
        dPayment = 0
        dDate = DTPFrom.Value
        dReceipt = IIf(rs!OpeningBalance >= 0, Val("" & rs!OpeningBalance), 0)
        dPayment = IIf(rs!OpeningBalance < 0, Val("" & rs!OpeningBalance), 0)
        MGrid.TextMatrix(MGrid.Rows - 1, gDate) = Format(DTPFrom.Value, "dd-MM-yyyy")
        MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Opening Balance"
        MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(Abs(dReceipt), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(Abs(dPayment), "0.00")
        
        
        rs.MoveFirst
    End If
    While rs.EOF = False
        If dDate <> DateValue("" & rs!TransactionDate) Then
            MGrid.AddItem ""
            MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Closing Balance"
            MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(IIf(dReceipt - dPayment <= 0, Abs(dReceipt - dPayment), 0), "0.00")
            MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(IIf(dReceipt - dPayment > 0, Abs(dReceipt - dPayment), 0), "0.00")
            MGrid.AddItem ""
            
            MGrid.AddItem ""
            MGrid.TextMatrix(MGrid.Rows - 1, gDate) = Format("" & rs!TransactionDate, "dd-MM-yyyy")
            MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Opening Balance"
            MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(IIf(dReceipt - dPayment >= 0, Abs(dReceipt - dPayment), 0), "0.00")
            MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(IIf(dReceipt - dPayment < 0, Abs(dReceipt - dPayment), 0), "0.00")
            dDate = DateValue("" & rs!TransactionDate)
            dReceipt = IIf(dReceipt - dPayment >= 0, Abs(dReceipt - dPayment), 0)
            dPayment = IIf(dReceipt - dPayment < 0, Abs(dReceipt - dPayment), 0)
        End If
        MGrid.AddItem ""
        dReceipt = dReceipt + Val("" & rs!Income)
        dPayment = dPayment + Val("" & rs!Expense)
        MGrid.TextMatrix(MGrid.Rows - 1, gToBy) = IIf(Trim("" & rs!Type) = "P", "To", "By")
        MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = rs!Type & rs!TransactionNo & " " & rs!AccountDescription & "," & rs!Narration
        MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(Abs(Val("" & rs!Income)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(Abs(Val("" & rs!Expense)), "0.00")
        
        rs.MoveNext
    Wend
    rs.Close
    MGrid.AddItem ""
    MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Total"
    MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(dReceipt & "", "0.00")
    MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(dPayment & "", "0.00")

    MGrid.AddItem ""
    MGrid.TextMatrix(MGrid.Rows - 1, gDescription) = "Closing Balance"
    MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(IIf(dReceipt - dPayment <= 0, Abs(dReceipt - dPayment), 0), "0.00")
    MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(IIf(dReceipt - dPayment > 0, Abs(dReceipt - dPayment), 0), "0.00")
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
  ' Create a new Excel worksheet...
    OLEExcel.CreateEmbed vbNullString, "Excel.Sheet"

  ' Now, pre-fill it with some data you
  ' can use. The OLE.Object property returns a
  ' workbook object, and you can use Sheets(1)
  ' to get the first sheet.
    lRowCount = MGrid.Rows
    lColCount = MGrid.Cols
    ReDim xData(1 To lRowCount + 1, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

  ' It is much more efficient to use an array to
  ' pass data to Excel than to push data over
  ' cell-by-cell, so you can use an array.

  ' Add some column headers to the array...
    xData(1, 1) = "Date"
    xData(1, 2) = " "
    xData(1, 3) = "Description"
    xData(1, 4) = "Receipt"
    xData(1, 5) = "Payment"

  ' Now add some data...
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i

  ' Assign the data to Excel...
    oExcelSheet.Range("A3:E" & lRowCount + 3).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Day Report From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")
    'oExcelSheet.Range("B9:E9").FormulaR1C1 = "=SUM(R[-5]C:R[-2]C)"

  ' Do some auto formatting...
    oExcelSheet.Range("A1:E" & lRowCount + 3).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next
    ' Delete the existing test file (if any)...
    Kill App.Path & "\Reports\DayBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

  ' Save the file as a native XLS file...
    oExcel.SaveAs App.Path & "\Reports\DayBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
  ' Close the OLE object and remove it...
    OLEExcel.Close
    OLEExcel.Delete
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\DayBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\DayBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
End Sub

Private Sub DTPTo_Change()
    MGrid.Rows = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase("Storage.mdb", False, False, "MS Access;PWD=12345abcde!")
    MGridInitialise
    DTPFrom.Value = Date
    DTPTo.Value = Date
End Sub

'Private Sub printReport()
'Dim sHeader(5) As String
'    sHeader(0) = "Date"
'    sHeader(1) = "To/By"
'    sHeader(2) = "Description"
'    sHeader(3) = "Receipt"
'    sHeader(4) = "Payment"
'    printGrid MGrid, sHeader, "Day Book as for " & Format(DTPTo.Value, "dd-MMM-yyyy")
'End Sub
