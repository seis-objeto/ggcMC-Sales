VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmInsUpsale 
   BorderStyle     =   0  'None
   Caption         =   "Accident Insurance Upsale"
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   5565
      TabIndex        =   23
      Top             =   600
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmInsUpsale.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5565
      TabIndex        =   24
      Top             =   1230
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmInsUpsale.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4425
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   7805
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrControl.xrFrame xrFrame4 
         Height          =   2220
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   2070
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   3916
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   7
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1785
            Width           =   2085
         End
         Begin VB.CheckBox chkField 
            Caption         =   "&Issue OR"
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Tag             =   "wt0;fb0"
            Top             =   405
            Width           =   1440
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   3375
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   105
            Width           =   1455
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   11
            Top             =   75
            Width           =   1455
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   8
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   18
            Top             =   1125
            Width           =   2085
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   9
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   20
            Top             =   1455
            Width           =   2085
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   16
            Top             =   795
            Width           =   3750
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
            Height          =   195
            Index           =   5
            Left            =   75
            TabIndex        =   21
            Top             =   1815
            Width           =   915
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account No"
            Height          =   195
            Index           =   9
            Left            =   75
            TabIndex        =   19
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amount"
            Height          =   195
            Index           =   4
            Left            =   2265
            TabIndex        =   12
            Top             =   150
            Width           =   1050
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PR No"
            Height          =   195
            Index           =   3
            Left            =   30
            TabIndex        =   10
            Top             =   120
            Width           =   480
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check No"
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   17
            Top             =   1170
            Width           =   720
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            Height          =   195
            Index           =   13
            Left            =   75
            TabIndex        =   15
            Top             =   825
            Width           =   840
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   525
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   1500
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   926
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3360
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   105
            Width           =   1455
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   7
            Top             =   105
            Width           =   1455
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Amount"
            Height          =   195
            Index           =   1
            Left            =   2340
            TabIndex        =   8
            Top             =   150
            Width           =   945
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OR No"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   6
            Top             =   150
            Width           =   495
         End
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   1365
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   90
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   2408
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   540
            Index           =   1
            Left            =   720
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   4125
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   405
            Width           =   4125
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   90
            Width           =   2235
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   4
            Top             =   720
            Width           =   570
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   2
            Top             =   405
            Width           =   420
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   0
            Top             =   75
            Width           =   345
         End
      End
   End
End
Attribute VB_Name = "frmInsUpsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmAdvancePayment"

Private p_oAppDrivr As clsAppDriver
Private p_oSkin As clsFormSkin
Private p_bIsOkey As Boolean
Private p_bLoaded As Boolean
Private p_sBankIDxx As String
Private p_nTranAmtx As Currency
Private p_bIssuedOR As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Let ClientName(Value As String)
   txtField(0).Text = Value
End Property

Property Let Address(Value As String)
   txtField(1).Text = Value
End Property

Property Let Transact(Value As String)
   txtField(2).Text = Format(Value, "Mmm DD, YYYY")
End Property

Property Let Amount(Value As Currency)
   p_nTranAmtx = Value
   txtField(3).Text = Format(p_nTranAmtx, "#,##0.00")
End Property

Property Let ORNo(Value As String)
   txtField(4).Text = Value
End Property

Property Get ORNo() As String
   If p_bIsOkey Then
      'If release OR was check then return the value from PR as the OR No.
      If chkField.Value = 1 Then
         ORNo = txtField(5).Text
      Else
         ORNo = txtField(4).Text
      End If
   End If
End Property

Property Get PRNo() As String
   If p_bIsOkey Then
      If chkField.Value = 0 Then
         PRNo = txtField(5).Text
      End If
   End If
End Property

Property Get BankID() As String
   If p_bIsOkey Then
      BankID = p_sBankIDxx
   End If
End Property

Property Get CheckNo() As String
   If p_bIsOkey Then
      CheckNo = txtField(8)
   End If
End Property

Property Get AccountNo() As String
   If p_bIsOkey Then
      AccountNo = txtField(9)
   End If
End Property

Property Get CheckDate() As String
   If p_bIsOkey Then
      CheckDate = txtField(7)
   End If
End Property

Property Get IsOkey() As Boolean
   IsOkey = p_bIsOkey
End Property

Property Get IsORIssued() As Boolean
   IsORIssued = p_bIssuedOR
End Property

Private Sub chkField_Click()
   If chkField.Value = 1 Then
      Label1(3).Caption = "OR No"
      p_bIssuedOR = True
   Else
      Label1(3).Caption = "PR No"
      p_bIssuedOR = False
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsProcName As String
   
   lsProcName = "cmdButton_Click"
   'On Error GoTo errProc
   
   Select Case Index
   Case 0
      If Trim(txtField(4).Text) <> "" And txtField(3) <> 0 Then
         p_bIsOkey = True
         Me.Hide
      Else
         If txtField(5).Text = "" Then
            If p_bIssuedOR Then
               MsgBox "Please enter a correct OR No!", vbYesNo, "Confirmation"
               GoTo endProc
            Else
               MsgBox "Please enter a correct PR No!", vbYesNo, "Confirmation"
               GoTo endProc
            End If
         End If

         If p_sBankIDxx = "" Then
            MsgBox "Please enter a Bank!", vbYesNo, "Confirmation"
            GoTo endProc
         End If

         If txtField(8) = "" Then
            MsgBox "Please enter the Check No!", vbYesNo, "Confirmation"
            GoTo endProc
         End If

         If txtField(9) = "" Then
            MsgBox "Please enter the Account No!", vbYesNo, "Confirmation"
            GoTo endProc
         End If

         If Not IsDate(txtField(7)) Then
            MsgBox "Please enter the Clearing Date!", vbYesNo, "Confirmation"
            GoTo endProc
         End If
         
         p_bIsOkey = True
         Me.Hide
      End If
         
'      If Trim(txtField(4).Text) = "" And txtField(3) > 0 Then
'         MsgBox "Please enter a correct OR No!", vbYesNo, "Confirmation"
'         GoTo endProc
'      Else
'         If txtField(5).Text = "" Then
'            If p_bIssuedOR Then
'               MsgBox "Please enter a correct OR No!", vbYesNo, "Confirmation"
'               GoTo endProc
'            Else
'               MsgBox "Please enter a correct PR No!", vbYesNo, "Confirmation"
'               GoTo endProc
'            End If
'         End If
'
'         If p_sBankIDxx = "" Then
'            MsgBox "Please enter a Bank!", vbYesNo, "Confirmation"
'            GoTo endProc
'         End If
'
'         If txtField(8) = "" Then
'            MsgBox "Please enter the Check No!", vbYesNo, "Confirmation"
'            GoTo endProc
'         End If
'
'         If txtField(9) = "" Then
'            MsgBox "Please enter the Account No!", vbYesNo, "Confirmation"
'            GoTo endProc
'         End If
'
'         If Not IsDate(txtField(7)) Then
'            MsgBox "Please enter the Clearing Date!", vbYesNo, "Confirmation"
'            GoTo endProc
'         End If
'      End If
   Case 1
      If MsgBox("Do you really want to cancel the entry?", vbYesNo, "Confirmation") Then
         p_bIsOkey = False
         Me.Hide
      End If
   End Select

endProc:
   Exit Sub
errProc:
   'ShowError lsProcName & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   If Not p_bLoaded Then
      If Trim(txtField(4).Text) = "" Then
         txtField(4).Text = GetNextOR()
      Else
         txtField(4).Text = GetNextOR(txtField(4).Text)
      End If
      
'      'Disable the PR part here
'      txtField(10).Enabled = False
'      txtField(8).Enabled = False
'      txtField(9).Enabled = False
'      txtField(7).Enabled = False
'      chkField.Enabled = False
      
      p_bLoaded = True
   End If
End Sub

Private Sub Form_Load()
   Set p_oSkin = New clsFormSkin
   With p_oSkin
      Set .AppDriver = p_oAppDrivr
      Set .Form = Me
      .DisableClose = True
      .ApplySkin xeFormTransDetail
   End With
End Sub

Private Function GetNextOR(Optional ByVal fsCode As String) As String
   If IsEmpty(fsCode) = True Then fsCode = IFNull(p_oAppDrivr.Config("sReceiptx"), "0")
   
   GetNextOR = Format(CDbl(IIf(fsCode = "", 0, fsCode)) + 1, String(Len(fsCode), "0"))
End Function

Public Sub getBanks(ByVal lsValue As String, ByVal lbExact As Boolean, ByVal lbByCode As Boolean)
   Dim lrs As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsOldProc As String
   Dim lsSearch As String
   Dim lsSQL As String
   
   lsOldProc = "getBanks"
   'On Error GoTo errProc
      
   If txtField(10).Tag = lsValue And Trim(lsValue) <> "" Then
      GoTo endProc
   End If
   
   lsSQL = "SELECT" _
               & "  sBankIDxx" _
               & ", sBankName" _
            & " FROM Banks" _
            & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
               & IIf(Not lbByCode _
               , IIf(Not lbExact, " AND sBankName LIKE " & strParm(lsValue & "%") _
               , " AND sBankName = " & strParm(lsValue)) _
               , " AND sBankIDxx = " & strParm(lsValue)) _
            & " ORDER BY sBankName"
   
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lrs.EOF Then
      txtField(10).Text = ""
      txtField(10).Tag = ""
      p_sBankIDxx = ""
      GoTo endProc
   End If
   
   If lrs.RecordCount = 1 Then
      txtField(10).Text = lrs("sBankName")
      txtField(10).Tag = lrs("sBankName")
      p_sBankIDxx = lrs("sBankIDxx")
   Else
      lsSearch = KwikBrowse(p_oAppDrivr, lrs _
                        , "sBankIDxx»sBankName" _
                        , "BankID»Bank Name" _
                        , "@»@")
      
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         txtField(10).Text = lsSelected(1)
         txtField(10).Tag = lsSelected(1)
         p_sBankIDxx = lsSelected(0)
      Else
         txtField(10).Text = ""
         txtField(10).Tag = ""
         p_sBankIDxx = ""
      End If
   End If
      
endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
'   ShowError lsOldProc & " ( " & lsValue & _
'                        ", " & lbExact & " ) "
   GoTo endProc
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 10 Then
      If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
         Call getBanks(txtField(10), False, False)
      End If
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 4  'OR
      'Does user entered a value for the OR No
      If Trim(txtField(4)) <> "" Then
         txtField(3).Text = Format(p_nTranAmtx, "#,##0.00")
         'Was PR No filled in before?
         If Trim(txtField(5)) <> "" Then
            If MsgBox("This will erase the entry for the PR part." & vbCrLf & _
                      "Do you want to continue?", vbYesNo, "MAFRE OR/PR Entry Confirmation") = vbYes Then
               txtField(5).Text = ""
               txtField(6).Text = "0.00"
               txtField(10).Text = ""
               txtField(8).Text = ""
               txtField(9).Text = ""
               txtField(7).Text = ""
               p_sBankIDxx = ""
               
               'Reset value for Is Release OR
               chkField.Value = 0
               Label1(3).Caption = "PR No"
            
'               'Disable the PR part here
'               txtField(10).Enabled = False
'               txtField(8).Enabled = False
'               txtField(9).Enabled = False
'               txtField(7).Enabled = False
'               chkField.Enabled = False
            Else
               txtField(4).Text = ""
               txtField(6).Text = Format(p_nTranAmtx, "#,##0.00")
            End If
         End If
      End If
   Case 5
      If Trim(txtField(5)) <> "" Then
         txtField(6).Text = Format(p_nTranAmtx, "#,##0.00")
         If CDbl(txtField(3).Text) > 0# Then 'And Trim(txtField(4).Text) <> ""
            If MsgBox("This will erase the entry for the OR part." & vbCrLf & _
                      "Do you want to continue?", vbYesNo, "MAFRE OR/PR Entry Confirmation") = vbYes Then
               txtField(3).Text = "0.00"
               txtField(4).Text = ""
               
'               'Enable the PR part here
'               txtField(10).Enabled = True
'               txtField(8).Enabled = True
'               txtField(9).Enabled = True
'               txtField(7).Enabled = True
'               chkField.Enabled = True
            Else
               txtField(5).Text = ""
               txtField(3).Text = Format(p_nTranAmtx, "#,##0.00")
            End If
         End If
      End If
   Case 10
      Call getBanks(txtField(10), True, False)
   End Select
   
End Sub
