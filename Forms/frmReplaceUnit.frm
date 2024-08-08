VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmReplaceUnit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5535
      TabIndex        =   13
      Top             =   1215
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
      Picture         =   "frmReplaceUnit.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   5535
      TabIndex        =   12
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
      Picture         =   "frmReplaceUnit.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2040
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3598
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1260
         TabIndex        =   9
         Top             =   1575
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1260
         TabIndex        =   7
         Top             =   1215
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         Top             =   855
         Width           =   3600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   495
         Width           =   3600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   1
         Top             =   135
         Width           =   3600
      End
      Begin VB.Label Label1 
         Caption         =   "Regis Amount"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1575
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Unit Price"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1215
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Model"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   855
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Frame No"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Engine No"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   135
         Width           =   930
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   675
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2625
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1191
      BackColor       =   12632256
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1365
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   165
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credited Amount"
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   10
         Top             =   210
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmReplaceUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oSkin As FormSkin

Private pnCredAmt As Double
Private pbCancel As Boolean

Dim lbSearch As Boolean

Property Let EngineNo(Value As String)
   txtField(0) = Value
End Property

Property Let FrameNo(Value As String)
   txtField(1) = Value
End Property

Property Let Model(Value As String)
   txtField(2) = Value
End Property

Property Let UnitPrice(Value As Double)
   txtField(3) = Value
End Property

Property Let RegisAmount(Value As Double)
   txtField(4) = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Property Get CreditAmount() As Double
   CreditAmount = CDbl(txtField(5))
End Property

Private Sub cmdButton_Click(Index As Integer)
   pbCancel = Index = 1
   Me.Hide
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me

   Set p_oSkin = New FormSkin
   Set p_oSkin.AppDriver = oApp
   Set p_oSkin.Form = Me
   p_oSkin.ApplySkin xeFormTransDetail
   
   txtField(5) = Format(CDbl(txtField(3)) - CDbl(txtField(4)), "#,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 5 Then
      If Not (CDbl(txtField(5)) = CDbl(txtField(3)) _
               Or CDbl(txtField(5)) = CDbl(txtField(3)) - CDbl(txtField(4))) Then
         txtField(5) = txtField(3)
      End If
      txtField(5) = Format(CDbl(txtField(5)), "#,##0.00")
   End If
End Sub
