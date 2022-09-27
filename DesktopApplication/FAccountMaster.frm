VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FAccountMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Master"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddGroup 
      Caption         =   "Add Group"
      Height          =   540
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Width           =   1890
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   540
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   540
      Left            =   6015
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CDeleteAccount 
      Caption         =   "Delete Account"
      Height          =   540
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CAddNew 
      Caption         =   "Add New"
      Height          =   540
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CFindNext 
      Caption         =   "Find Next"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5085
      Width           =   1470
   End
   Begin MSComctlLib.TreeView TrAccounts 
      Height          =   4785
      Left            =   285
      TabIndex        =   0
      Top             =   210
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   8440
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.TextBox TAddress3 
      Height          =   345
      Left            =   7140
      TabIndex        =   5
      Top             =   2055
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TAddress2 
      Height          =   345
      Left            =   7140
      TabIndex        =   4
      Top             =   1620
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TNarration 
      Height          =   345
      Left            =   7140
      TabIndex        =   7
      Top             =   2925
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   4920
      TabIndex        =   21
      Top             =   2970
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   345
      Left            =   7140
      TabIndex        =   2
      Top             =   750
      Width           =   2520
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4445;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoStatus 
      Height          =   330
      Left            =   7140
      TabIndex        =   8
      Top             =   3360
      Width           =   2520
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4445;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   405
      Left            =   4920
      TabIndex        =   20
      Top             =   3375
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Status"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   405
      Left            =   4920
      TabIndex        =   19
      Top             =   2535
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Phone"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPhone 
      Height          =   345
      Left            =   7140
      TabIndex        =   6
      Top             =   2490
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   405
      Left            =   4920
      TabIndex        =   18
      Top             =   1215
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Address"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress1 
      Height          =   345
      Left            =   7140
      TabIndex        =   3
      Top             =   1185
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   4920
      TabIndex        =   17
      Top             =   750
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Name"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   405
      Left            =   4920
      TabIndex        =   16
      Top             =   300
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAccountCode 
      Height          =   330
      Left            =   7140
      TabIndex        =   1
      Top             =   330
      Width           =   2520
      VariousPropertyBits=   746604575
      Size            =   "4445;582"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TFind 
      Height          =   315
      Left            =   270
      TabIndex        =   11
      Top             =   5100
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FAccountMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCreateNewGroup As Boolean

Private Sub getAccount()
Dim rs As Recordset
    
    CoAccount.Clear
    
    Set rs = db.OpenRecordset("Select AccountMaster.AccountName From AccountMaster  Order By AccountMaster.AccountName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    While rs.EOF = False
        CoAccount.AddItem "" & rs!AccountName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CAddNew_Click()
    
    If (TrAccounts.Nodes.Count = 0) Then
        MsgBox "Please Create a a Group First !", vbInformation
        Exit Sub
    End If
    
    If Left(Trim(TrAccounts.SelectedItem.Key), 1) = "B" Then
        MsgBox "Please Select any Group to create Account !", vbInformation
        Exit Sub
    End If

    clearEditControls
    enableDisableControlsOnAdd
    TAccountCode = getNewAccountCode
    CoAccount.SetFocus
End Sub

Private Sub CAddGroup_Click()
    clearEditControls
    enableDisableControlsOnGroup
    TAccountCode = getNewAccountCode
    CoAccount.SetFocus
    bCreateNewGroup = True
End Sub

Private Sub enableDisableControlsOnGroup()
    TAccountCode.Enabled = False
    CoAccount.Enabled = True
    TAddress1.Enabled = False
    TAddress2.Enabled = False
    TAddress3.Enabled = False
    TPhone.Enabled = False
    TNarration.Enabled = False
    CoStatus.Enabled = True
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDeleteAccount_Click()
Dim rs As Recordset
    
    If Trim(TAccountCode.Text) = "" Then
        MsgBox "Please Select Any Account to Delete !", vbInformation
        Exit Sub
    End If
        
    If checkForChildAccounts(Trim(TAccountCode.Text)) Then
        MsgBox "The Group has Account Items, Please Delete them First !", vbInformation
        Exit Sub
    End If
        
    Set rs = db.OpenRecordset("Select AccountMaster.* From AccountMaster Where (AccountMaster.Code = '" & Trim(TAccountCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "The Account doesnt Exist !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted the Account !", vbInformation
    
    refreshTree
    clearEditControls
End Sub

Private Function checkForChildAccounts(sAccountCode As String) As Boolean
Dim rs As Recordset, bFound As Boolean

    Set rs = db.OpenRecordset("Select AccountMaster.* From AccountMaster Where (AccountMaster.GroupCode = '" & Trim(sAccountCode) & "' )")
    If rs.RecordCount > 0 Then
        bFound = True
    Else
        bFound = False
    End If
    rs.Close
    
    checkForChildAccounts = bFound
End Function

Private Sub CFindNext_Click()
Static lFindIndex As Long
Static sFindWord As String
    
    If Trim(TFind.Text) <> sFindWord Then
        lFindIndex = 1
    Else
        lFindIndex = lFindIndex + 1
    End If
    
    sFindWord = Trim(TFind.Text)
    
    Do While lFindIndex <= TrAccounts.Nodes.Count
        
        If InStr(1, LCase(TrAccounts.Nodes.Item(lFindIndex)), LCase(sFindWord), vbTextCompare) > 0 Then
            TrAccounts.Nodes.Item(lFindIndex).Selected = True
            getDetailsOfAccount
            TrAccounts.SetFocus
            Exit Do
        End If
        lFindIndex = lFindIndex + 1
    Loop
    
    If lFindIndex > TrAccounts.Nodes.Count Then
        MsgBox "No more Items !", vbInformation
        lFindIndex = 1
        Exit Sub
    End If
End Sub

Private Sub CoStatus_GotFocus()
    CoStatus.SelStart = 0
    CoStatus.SelLength = Len(CoStatus.Text)
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sAccountCode As String, sParenttype As String
Dim sParentCode As String

    If Trim(TAccountCode.Text) = "" Then
        MsgBox "Please Select a Account to Edit or click Add New button To add new Account", vbInformation
        Exit Sub
    ElseIf Trim(CoAccount.Text) = "" Then
        MsgBox "Please Enter needed Informations !", vbInformation
        CoAccount.SetFocus
        Exit Sub
    End If
    
    'Determines GroupCode
    If TrAccounts.Nodes.Count > 0 Then
        If (bCreateNewGroup) Then
            sParenttype = ""
            sParentCode = ""
        Else
            sParenttype = Trim(Left(TrAccounts.SelectedItem.Key, 1))
            sParentCode = Trim(Right(TrAccounts.SelectedItem.Key, Len(TrAccounts.SelectedItem.Key) - 1))
        End If
    Else
        sParenttype = ""
        sParentCode = ""
    End If

    Set rs = db.OpenRecordset("Select AccountMaster.* From AccountMaster Where (AccountMaster.Code = '" & Trim(TAccountCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
    Else
        sStatus = "Added"
        TAccountCode.Text = getNewAccountCode()
        rs.AddNew
        rs!Code = Trim(TAccountCode.Text)
        rs!Type = IIf(Trim(sParenttype) = "", "AGroup", "BAccount")
        rs!GroupCode = sParentCode
    End If
    rs!AccountName = Trim(CoAccount.Text)
    rs!Address1 = Trim(TAddress1.Text)
    rs!Address2 = Trim(TAddress2.Text)
    rs!Address3 = Trim(TAddress3.Text)
    rs!Phone = Trim(TPhone.Text)
    rs!Narration = Trim(TNarration.Text)
    rs!Status = IIf((CoStatus.ListIndex = 0), True, False)
    rs.Update
    rs.Close
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    clearEditControls
    bCreateNewGroup = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF And ((Shift And 7) = 2)) Then
        CFindNext_Click
    ElseIf (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteAccount_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    CoStatus.AddItem "Enabled"
    CoStatus.AddItem "Disabled"
    
    refreshTree
    enableDisableControls
    getAccount
    bCreateNewGroup = False
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrAccounts.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName,AccountMaster.Type,AccountMaster.GroupCode From AccountMaster Order By AccountMaster.Type,AccountMaster.AccountName")
    While rs.EOF = False
        If Trim(rs!Type) = "AGroup" Then
            TrAccounts.Nodes.Add , , "A" & rs!Code, rs!AccountName
            'TrAccounts.Nodes(TrAccounts.Nodes.Count).Bold = True
            'TrAccounts.Nodes(TrAccounts.Nodes.Count).ForeColor = &H808080
        ElseIf Trim(rs!Type) = "BAccount" Then
            TrAccounts.Nodes.Add "A" & rs!GroupCode, tvwChild, "B" & rs!Code, rs!AccountName
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TAccountCode_GotFocus()
    TAccountCode.SelStart = 0
    TAccountCode.SelLength = Len(TAccountCode.Text)
End Sub


Private Sub CoAccount_GotFocus()
    CoAccount.SelStart = 0
    CoAccount.SelLength = Len(CoAccount.Text)
End Sub

Private Sub TFind_GotFocus()
    TFind.SelStart = 0
    TFind.SelLength = Len(TFind.Text)
End Sub

Private Sub TrAccounts_Click()
    enableDisableControls
End Sub

Private Sub TrAccounts_NodeClick(ByVal Node As MSComctlLib.Node)
    enableDisableControls
    If TrAccounts.Nodes.Count > 0 Then
        getDetailsOfAccount
    End If
End Sub

Private Sub getDetailsOfAccount()
Dim rs As Recordset
        
    Set rs = db.OpenRecordset("Select AccountMaster.* From AccountMaster Where (AccountMaster.Code = '" & Trim(Right(TrAccounts.SelectedItem.Key, Len(TrAccounts.SelectedItem.Key) - 1)) & "' )")
        
    If rs.RecordCount > 0 Then
        
        TAccountCode.Text = "" & rs!Code
        CoAccount.Text = "" & rs!AccountName
        TAddress1.Text = "" & rs!Address1
        TAddress2.Text = "" & rs!Address2
        TAddress3.Text = "" & rs!Address3
        TPhone.Text = "" & rs!Phone
        TNarration.Text = "" & rs!Narration
        CoStatus.ListIndex = IIf((rs!Status = True), 0, 1)
    Else
        clearEditControls
    End If
    rs.Close
End Sub

Private Sub enableDisableControlsOnAdd()
    
    If Left(TrAccounts.SelectedItem.Key, 1) = "A" Then
        
        TAccountCode.Enabled = False
        CoAccount.Enabled = True
        TAddress1.Enabled = True
        TAddress2.Enabled = True
        TAddress3.Enabled = True
        TPhone.Enabled = True
        TNarration.Enabled = True
        CoStatus.Enabled = True
    ElseIf Left(TrAccounts.SelectedItem.Key, 1) = "B" Then
        
        TAccountCode.Enabled = False
        CoAccount.Enabled = True
        TAddress1.Enabled = True
        TAddress2.Enabled = True
        TAddress3.Enabled = True
        TPhone.Enabled = True
        TNarration.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub enableDisableControls()
    
    If TrAccounts.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Left(TrAccounts.SelectedItem.Key, 1) = "A" Then
        
        TAccountCode.Enabled = False
        CoAccount.Enabled = True
        TAddress1.Enabled = False
        TAddress2.Enabled = False
        TAddress3.Enabled = False
        TPhone.Enabled = False
        TNarration.Enabled = False
        CoStatus.Enabled = True
    ElseIf Left(TrAccounts.SelectedItem.Key, 1) = "B" Then
                
        TAccountCode.Enabled = False
        CoAccount.Enabled = True
        TAddress1.Enabled = True
        TAddress2.Enabled = True
        TAddress3.Enabled = True
        TPhone.Enabled = True
        TNarration.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub clearEditControls()
    TAccountCode.Text = ""
    CoAccount.Text = ""
    TAddress1.Text = ""
    TAddress2.Text = ""
    TAddress3.Text = ""
    TPhone.Text = ""
    TNarration.Text = ""
    CoStatus.Text = ""
End Sub

Private Function getParentAccount(sAccountCode As String) As String
Dim rs As Recordset, sParentCode As String
    
    Set rs = db.OpenRecordset("Select AccountMaster.GroupCode From AccountMaster Where (AccountMaster.Code = '" & Trim(sAccountCode) & "' )")
    If rs.RecordCount > 0 Then
        sParentCode = "" & rs!GroupCode
    Else
        sParentCode = ""
    End If
    rs.Close
    
    getParentAccount = sParentCode
End Function

