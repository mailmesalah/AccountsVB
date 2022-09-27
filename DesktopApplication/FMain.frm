VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FMain 
   Caption         =   "DeXtop - Accounts Software"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   765
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2535
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSComDlg.CommonDialog CoDialog 
      Left            =   945
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MMasters 
      Caption         =   "Masters"
      Begin VB.Menu MMAccountMaster 
         Caption         =   "Account Master"
      End
   End
   Begin VB.Menu MTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu MTDayTransaction 
         Caption         =   "Day Transaction"
      End
      Begin VB.Menu mb 
         Caption         =   "-"
      End
      Begin VB.Menu MTReceipt 
         Caption         =   "Receipt"
      End
      Begin VB.Menu MTPayment 
         Caption         =   "Payment"
      End
      Begin VB.Menu Mc 
         Caption         =   "-"
      End
      Begin VB.Menu MTJournalVoucher 
         Caption         =   "Journal Voucher"
      End
   End
   Begin VB.Menu MReports 
      Caption         =   "Reports"
      Begin VB.Menu MRDayBook 
         Caption         =   "Day Book"
      End
      Begin VB.Menu MRLedger 
         Caption         =   "Ledger"
      End
   End
   Begin VB.Menu MSettings 
      Caption         =   "Settings"
      Begin VB.Menu MSBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu MSRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mm 
         Caption         =   "-"
      End
      Begin VB.Menu MSAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub backUp()
On Error GoTo GoOut
Dim x As Long
    
    'BACKUP DATA
    Dim fso As Object, s As String
    
    CoDialog.CancelError = True
    
    CoDialog.FileName = "DeXtop_" & Day(Date) & "_" & Month(Date) & "_" & Year(Date)
    CoDialog.Filter = "mdb"
    CoDialog.ShowSave
    Set fso = CreateObject("Scripting.FileSystemObject")
    x = fso.CopyFile(App.Path & "/Storage.mdb", CoDialog.FileName & ".mdb", True)
    
    x = MsgBox("Successfully Exported !", vbInformation)
    Exit Sub
GoOut:
    x = MsgBox("Backup was Failed : " & Err.Description, vbInformation)
End Sub

Private Sub reStore()
On Error GoTo GoOut
Dim x As Long
    
    If (MsgBox("Are you sure to Restore ? ,Current Data will be Overwritten !", vbDefaultButton2 Or vbYesNo) = vbNo) Then
        Exit Sub
    End If
    
    'RESTORE DATA
    Dim fso As Object
    
    CoDialog.CancelError = True
    CoDialog.Filter = "mdb"
    CoDialog.ShowOpen
    Set fso = CreateObject("Scripting.FileSystemObject")
    x = fso.CopyFile(CoDialog.FileName, App.Path & "/Storage.mdb", True)
    
    x = MsgBox("Successfully Restored !", vbInformation)
    Exit Sub
    
GoOut:
    x = MsgBox("Restore was Failed : " & Err.Description, vbInformation)
End Sub

Private Sub Form_Load()
    initialisePublicVariables
End Sub

Private Sub MMAccountMaster_Click()
    FAccountMaster.Show
End Sub

Private Sub MRDayBook_Click()
    FDayBook.Show
End Sub

Private Sub MRLedger_Click()
    FLedger.Show
End Sub

Private Sub MTDayTransaction_Click()
    FDayTransaction.Show
End Sub

Private Sub MTJournalVoucher_Click()
    FJournalVoucher.Show
End Sub

Private Sub MTPayment_Click()
    FPayment.Show
End Sub

Private Sub MTReceipt_Click()
    FReceipt.Show
End Sub
