VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMC_Agent 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "MC Agent"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   7785
      TabIndex        =   23
      Top             =   1185
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
      Picture         =   "frmMC_Agent.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   7785
      TabIndex        =   22
      Top             =   555
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
      Picture         =   "frmMC_Agent.frx":077A
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3570
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   6297
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkGCard 
         Caption         =   "G-Card Holder"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6015
         TabIndex        =   24
         Tag             =   "wt0;fb0"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   5310
         MaxLength       =   15
         TabIndex        =   6
         Top             =   2820
         Width           =   2040
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2490
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1830
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1470
         MaxLength       =   128
         TabIndex        =   3
         Top             =   2160
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   2
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1290
         Width           =   5880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1470
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   5880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   5310
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1830
         Width           =   2040
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2820
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   7
         Top             =   3150
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Since"
         Height          =   195
         Index           =   12
         Left            =   3990
         TabIndex        =   20
         Top             =   2880
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   20
         Left            =   75
         TabIndex        =   13
         Top             =   1545
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   19
         Left            =   75
         TabIndex        =   12
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   10
         Top             =   690
         Width           =   510
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1575
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship"
         Height          =   195
         Index           =   17
         Left            =   75
         TabIndex        =   18
         Top             =   2550
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Dependent"
         Height          =   195
         Index           =   15
         Left            =   75
         TabIndex        =   15
         Top             =   1890
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee"
         Height          =   195
         Index           =   14
         Left            =   75
         TabIndex        =   17
         Top             =   2220
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Child"
         Height          =   195
         Index           =   13
         Left            =   3990
         TabIndex        =   16
         Top             =   1890
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission"
         Height          =   195
         Index           =   11
         Left            =   75
         TabIndex        =   19
         Top             =   2880
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Info"
         Height          =   195
         Index           =   10
         Left            =   75
         TabIndex        =   21
         Top             =   3210
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmMC_Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMC_Agent"

Private p_oAppDriver As clsAppDriver
Private pbCancelled As Boolean
Private psEmployCd As String
Private psRelatnCd As String

Private oSkin As clsFormSkin

Property Set AppDriver(oValue As clsAppDriver)
   Set p_oAppDriver = oValue
End Property
Property Let AgentID(sValue As String)
   txtOthers(0).Text = sValue
End Property

Property Let AgentName(sValue As String)
   txtOthers(1).Text = sValue
End Property

Property Let GCardHolder(bValue As Boolean)
   chkGCard.Value = IIf(bValue = True, 1, 0)
End Property

Property Let AgentAddress(sValue As String)
   txtOthers(2).Text = sValue
End Property

Property Get NoOfDependent() As Integer
   NoOfDependent = txtField(2).Text
End Property

Property Let NoOfDependent(sValue As Integer)
   txtField(2).Text = sValue
End Property

Property Get NoOfChild() As Integer
   NoOfChild = txtField(3).Text
End Property

Property Let NoOfChild(sValue As Integer)
   txtField(3).Text = sValue
End Property

Property Get Employee() As String
   Employee = txtField(4).Text
End Property

Property Let Employee(sValue As String)
   txtField(4).Text = Trim(sValue)
   txtField(4).Tag = Trim(sValue)
End Property

Property Get EmployeeCd() As String
   EmployeeCd = psEmployCd
End Property

Property Let EmployeeCd(sValue As String)
   Employee = getDescription(sValue, "Client_Master")
   psEmployCd = sValue
End Property

Property Get Relation() As String
   Relation = txtField(5).Text
End Property

Property Let Relation(sValue As String)
   txtField(5).Text = Trim(sValue)
   txtField(5).Tag = Trim(sValue)
End Property

Property Get RelationCd() As String
   RelationCd = psRelatnCd
End Property

Property Let RelationCd(sValue As String)
   Relation = getDescription(sValue, "Relation")
   psRelatnCd = sValue
End Property

Property Get Commission() As Double
   Commission = CDbl(txtField(6).Text)
End Property

Property Let Commission(sValue As Double)
   txtField(6).Text = Format(sValue, "#,##0.00")
End Property

Property Get ClientSince() As Date
   Commission = CDate(txtField(7).Text)
End Property

Property Let ClientSince(sValue As Date)
   txtField(7).Text = Format(sValue, "MMM DD, YYYY")
End Property

Property Get OtherInfo() As String
   OtherInfo = txtField(8).Text
End Property

Property Let OtherInfo(sValue As String)
   txtField(8).Text = sValue
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   pbCancelled = False
   Select Case Index
   Case 0
      Me.Hide
   Case 1
      pbCancelled = True
      Me.Hide
   End Select
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
   Case vbKeyReturn, vbKeyDown
      SetNextFocus
   Case vbKeyUp
      SetPreviousFocus
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDriver
   Set oSkin.Form = Me
   
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
   
   txtField(1).Text = p_oAppDriver.BranchName
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Function getEmployee(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getRelation"
   'On Error GoTo errProc
   
   If lsValue <> "" Then
      If lsValue = txtField(4).Tag Then GoTo endProc
      
      If lbSearch Then
         lsSQL = "a.sCompnyNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSQL = "a.sCompnyNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endProc
   End If
   
   lsSQL = "SELECT a.sClientID" & _
               ", a.sCompnyNm" & _
               ", c.sBranchNm" & _
            " FROM Client_Master a" & _
               ", Employee_Master001 b" & _
               ", Branch c" & _
            " WHERE a.sClientID = b.sEmployID" & _
               " AND b.sBranchCd = c.sBranchCd" & _
               " AND b.cRecdStat = " & strParm(xeRecStateActive) & _
               IIf(lsSQL = "", "", " AND " & lsSQL) & _
            " ORDER BY a.sCompnyNm"
Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDriver.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then
      If Not lbSearch Then
         psEmployCd = ""
         txtField(4).Text = ""
         txtField(4).Tag = ""
      End If
      GoTo endProc
   End If
   
   If loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDriver, loRS _
                           , "sClientID»sCompnyNm»sBranchNm" _
                           , "ID»Employee Name»Branch")
      If lsSQL = "" Then
         psEmployCd = ""
         txtField(4).Text = ""
         txtField(4).Tag = ""
         GoTo endProc
      Else
         lasMaster = Split(lsSQL, "»")
         
         loRS.MoveFirst
         loRS.Find "sClientID = " & strParm(lasMaster(0)), 0, adSearchForward
      End If
   End If
      
   psEmployCd = loRS("sClientID")
   txtField(4).Text = loRS("sCompnyNm")
   txtField(4).Tag = loRS("sCompnyNm")
   
   getEmployee = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lbSearch & " )"
End Function

Private Function getRelation(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getRelation"
   'On Error GoTo errProc
   
   If lsValue <> "" Then
      If lsValue = txtField(5).Tag Then GoTo endProc
      
      If lbSearch Then
         lsSQL = "sRelatnDs LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSQL = "sRelatnDs = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endProc
   End If
   
   lsSQL = "SELECT sRelatnID" & _
               ", sRelatnDs" & _
            " FROM Relation" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
               IIf(lsSQL = "", "", " AND " & lsSQL) & _
            " ORDER BY sRelatnDs"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDriver.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then
      If Not lbSearch Then
         psRelatnCd = ""
         txtField(5).Text = ""
         txtField(5).Tag = ""
      End If
      GoTo endProc
   End If
   
   If loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDriver, loRS _
                           , "sRelatnID»sRelatnDs" _
                           , "ID»Relation")
      If lsSQL = "" Then
         psRelatnCd = ""
         txtField(5).Text = ""
         txtField(5).Tag = ""
         GoTo endProc
      Else
         lasMaster = Split(lsSQL, "»")
         
         loRS.MoveFirst
         loRS.Find "sRelatnID = " & strParm(lasMaster(0)), 0, adSearchForward
      End If
   End If
      
   psRelatnCd = loRS("sRelatnID")
   txtField(5).Text = loRS("sRelatnDs")
   txtField(5).Tag = loRS("sRelatnDs")
   
   getRelation = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lbSearch & " )"
End Function

Private Function getDescription(ByVal lsValue As String, ByVal lsTable As String) As String
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsProcName As String

   lsProcName = "getDescription"
   'On Error GoTo errProc
   
   If lsValue = "" Then GoTo endProc
   Select Case lsTable
   Case "Relation"
      lsSQL = "SELECT sRelatnDs" & _
            " FROM Relation" & _
            " WHERE sRelatnID = " & strParm(lsValue)
   Case "Client_Master"
      lsSQL = "SELECT sCompnyNm" & _
               " FROM Client_Master" & _
               " WHERE sClientID = " & strParm(lsValue)
   End Select
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDriver.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then
      GoTo endProc
   End If
   
   getDescription = loRS(0)
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lsTable & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With p_oAppDriver
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsProcName As String
   
   lsProcName = "txtField_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 4
            If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
               Call getEmployee(.Text, True)
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  Call getEmployee(.Text, True)
               End If
            End If
         Case 5
            If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
               Call getRelation(.Text, True)
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  Call getRelation(.Text, True)
               End If
            End If
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )"
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 6 Then
      txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
   End If
End Sub
