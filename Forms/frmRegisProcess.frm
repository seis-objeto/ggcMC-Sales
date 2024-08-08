VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmRegisProcess 
   BorderStyle     =   0  'None
   Caption         =   "Form of Registration"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2670
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4710
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.OptionButton optcRegisFrm 
         Caption         =   "Renewal"
         Height          =   345
         Index           =   3
         Left            =   3675
         TabIndex        =   7
         Top             =   2025
         Width           =   2340
      End
      Begin VB.OptionButton optcRegisFrm 
         Caption         =   "Transfer"
         Height          =   345
         Index           =   2
         Left            =   3675
         TabIndex        =   6
         Top             =   1680
         Width           =   2340
      End
      Begin VB.OptionButton optcRegisFrm 
         Caption         =   "Transfer and Renewal"
         Height          =   345
         Index           =   1
         Left            =   3675
         TabIndex        =   5
         Top             =   1335
         Width           =   2340
      End
      Begin VB.OptionButton optcRegisFrm 
         Caption         =   "New"
         Height          =   345
         Index           =   0
         Left            =   3675
         TabIndex        =   4
         Top             =   990
         Width           =   2340
      End
      Begin VB.Label FrameNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1455
         TabIndex        =   12
         Top             =   1605
         Width           =   1935
      End
      Begin VB.Label EngineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1455
         TabIndex        =   11
         Top             =   1290
         Width           =   1935
      End
      Begin VB.Label ModelName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1455
         TabIndex        =   10
         Top             =   975
         Width           =   1935
      End
      Begin VB.Label TransNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1590
         TabIndex        =   9
         Top             =   105
         Width           =   1920
      End
      Begin VB.Label Label2 
         Caption         =   "Form of Registration to Process"
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
         Left            =   3705
         TabIndex        =   8
         Top             =   750
         Width           =   2835
      End
      Begin VB.Shape Shape3 
         Height          =   1695
         Left            =   3570
         Top             =   810
         Width           =   3345
      End
      Begin VB.Shape Shape2 
         Height          =   1695
         Left            =   135
         Top             =   810
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame No."
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   3
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No."
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   2
         Top             =   1290
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Type"
         Height          =   195
         Index           =   10
         Left            =   255
         TabIndex        =   1
         Top             =   990
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   150
         Width           =   1350
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1650
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1935
      End
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   7470
      TabIndex        =   13
      Top             =   585
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
      Picture         =   "frmRegisProcess.frx":0000
   End
End
Attribute VB_Name = "frmRegisProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private p_oAppDrivr As clsAppDriver
Private oSkin As clsFormSkin

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Function getRegisForm() As Byte
   Dim lnCtr As Byte
   getRegisForm = 5
   
   For lnCtr = 0 To optcRegisFrm.Count - 1
      If optcRegisFrm(lnCtr).Value = True Then getRegisForm = lnCtr
   Next
End Function

Sub setRegisForm(ByVal Index As Byte)
   optcRegisFrm(Index).Value = True
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Me.Hide
End Sub

Private Sub Form_Activate()
   Dim lnRegisFrm As Byte
   lnRegisFrm = getRegisForm
   Call setRegisForm(IIf(lnRegisFrm = 5, 1, lnRegisFrm))
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
End Sub
