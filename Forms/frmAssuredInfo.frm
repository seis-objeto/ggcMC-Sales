VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmAssuredInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Assured Info"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   7815
      TabIndex        =   16
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
      Picture         =   "frmAssuredInfo.frx":0000
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2580
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   4551
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   5625
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1785
         Width           =   1740
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         Top             =   2115
         Width           =   2640
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   6390
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1455
         Width           =   975
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1455
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   915
         Width           =   6285
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Text            =   "Cuison, Michael Torres"
         Top             =   585
         Width           =   6285
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1785
         Width           =   2640
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   90
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         Height          =   195
         Index           =   4
         Left            =   4800
         TabIndex        =   12
         Top             =   1860
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   2175
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   6045
         TabIndex        =   8
         Top             =   1530
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   20
         Left            =   75
         TabIndex        =   4
         Top             =   1170
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   19
         Left            =   75
         TabIndex        =   2
         Top             =   585
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item No."
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
         TabIndex        =   0
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
         Height          =   195
         Index           =   15
         Left            =   75
         TabIndex        =   6
         Top             =   1515
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         Height          =   195
         Index           =   11
         Left            =   75
         TabIndex        =   10
         Top             =   1845
         Width           =   825
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2595
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   3135
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   4577
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1260
         TabIndex        =   34
         Text            =   "Cuison, Janine Kathleen Siquico"
         Top             =   2160
         Width           =   3030
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   4995
         TabIndex        =   36
         Text            =   "Wife"
         Top             =   2160
         Width           =   1380
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1260
         TabIndex        =   29
         Text            =   "Cuison, Janine Kathleen Siquico"
         Top             =   1830
         Width           =   3030
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   4995
         TabIndex        =   31
         Text            =   "Wife"
         Top             =   1830
         Width           =   1380
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4995
         TabIndex        =   26
         Text            =   "Wife"
         Top             =   1500
         Width           =   1380
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   24
         Text            =   "Cuison, Janine Kathleen Siquico"
         Top             =   1500
         Width           =   3030
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   675
         Index           =   1
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   420
         Width           =   6105
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   18
         Top             =   90
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3."
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
         Index           =   18
         Left            =   540
         TabIndex        =   32
         Top             =   2220
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   17
         Left            =   765
         TabIndex        =   33
         Top             =   2220
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         Height          =   195
         Index           =   14
         Left            =   4335
         TabIndex        =   35
         Top             =   2220
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2."
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
         Index           =   13
         Left            =   540
         TabIndex        =   27
         Top             =   1890
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   12
         Left            =   765
         TabIndex        =   28
         Top             =   1890
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         Height          =   195
         Index           =   10
         Left            =   4335
         TabIndex        =   30
         Top             =   1890
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1."
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
         Index           =   9
         Left            =   540
         TabIndex        =   22
         Top             =   1560
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         Height          =   195
         Index           =   8
         Left            =   4335
         TabIndex        =   25
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   7
         Left            =   765
         TabIndex        =   23
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Info"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Type"
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   17
         Top             =   150
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Benefeciaries"
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
         Index           =   2
         Left            =   45
         TabIndex        =   21
         Top             =   1200
         Width           =   1170
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   7815
      TabIndex        =   37
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAssuredInfo.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   7815
      TabIndex        =   38
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "C&lose"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAssuredInfo.frx":0EF4
   End
End
Attribute VB_Name = "frmAssuredInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMC_Agent"
Private Const pxeDEFAULTINSURANCE = "0001"   '30,000 coverage(FREE)
Private Const pxeSOURCECODE = "MCSl"

Private p_oAppDriver As clsAppDriver
Private WithEvents oTrans As clsAccidentInsurance
Attribute oTrans.VB_VarHelpID = -1
Private p_oClient As clsNeoClient 'insured object

Private oSkin As clsFormSkin

Private psParentxx As String
Private psClientID As String
Private psTransNox As String
Private psSerialID As String
Private pbShowMsg As Boolean
Private pbViewOnly As Boolean
Private pbCancelled As Boolean
Private pdTransact As Date
Private p_sORNoxxxx As String

Property Set AppDriver(oValue As clsAppDriver)
   Set p_oAppDriver = oValue
End Property

Property Let TransNox(ByVal Value As String)
   psTransNox = Value
End Property

Property Let ORNo(ByVal Value As String)
   p_sORNoxxxx = Value
End Property

Property Let ClientID(ByVal Value As String)
   psClientID = Value
End Property

Property Let MCSerial(ByVal Value As String)
   psSerialID = Value
End Property

Property Let SaleDate(ByVal Value As Date)
   pdTransact = Value
End Property

Property Let Parent(ByVal Value As String)
   psParentxx = Value
End Property

Property Let ViewOnly(ByVal Value As String)
   pbViewOnly = Value
End Property

Property Let showMessage(ByVal Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Cancelled()
   Cancelled = pbCancelled
End Property

Property Get Insurance()
   Set Insurance = oTrans
End Property

Property Set Client(Value As clsNeoClient)
   Set p_oClient = Value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   pbCancelled = False
   Select Case Index
   Case 0
      pbCancelled = False
      Me.Hide
   Case 1
      If MsgBox("Are you sure to disregard MAPFRE Insurance?", vbQuestion + vbYesNo, "") = vbYes Then
         pbCancelled = True
         Me.Hide
      End If
   Case 4
   Case 5
      If oTrans.UpdateTransaction Then
         xrFrame1.Enabled = True
         If Trim(oTrans.Master("sClientID")) = "" Then
            txtOthers(0).Locked = False
            txtOthers(1).Locked = False
            xrFrame2.Enabled = True
            txtField(1).SetFocus
         Else
            txtOthers(0).Locked = True
            txtOthers(1).Locked = True
            xrFrame2.Enabled = False
            txtDetail(0).SetFocus
         End If
         
         cmdButton(5).Visible = False
      End If
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
   Call InitFields

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDriver
   Set oSkin.Form = Me
   
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
   
   Set oTrans = New clsAccidentInsurance
   Set oTrans.AppDriver = p_oAppDriver
   oTrans.showMessage = pbShowMsg
   oTrans.ORNo = p_sORNoxxxx
   oTrans.InitTransaction
   oTrans.NewTransaction
      
   Call InitFields
   Call InitValue
endProc:
   Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitValue()
   Dim lotxt As TextBox
   Dim lnCtr As Integer
   
   With oTrans
      If Not pbViewOnly Then
         xrFrame1.Enabled = True
         xrFrame2.Enabled = True
      
         '.Master("sClientID") = psClientID
         Set .Client = p_oClient
                  
         txtField(0) = p_oClient.Master("sClientID")
         txtField(1) = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
         txtField(1).Tag = txtField(1).Text
         txtField(2) = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")
         txtField(3) = Format(p_oClient.Master("dBirthDte"), "Mmm dd, yyyy")
         txtField(4) = DateDiff("yyyy", p_oClient.Master("dBirthDte"), p_oAppDriver.SysDate)
         txtField(6) = IFNull(p_oClient.Master("sOccptnNm"), "N-O-N-E")
         txtField(7) = p_oClient.Master("sMobileNo")
         
         Select Case p_oClient.Master("cCvilStat")
         Case "0"
            txtField(5) = "SINGLE"
         Case "1"
            txtField(5) = "MARRIED"
         Case "2"
            txtField(5) = "SEPARATED"
         Case "3"
            txtField(5) = "WIDOWED"
         End Select
         
         If Len(psTransNox) > 12 Then
            If Mid(psTransNox, 13) = "TLM" Then
               .SearchMaster "sInsPrmID", "0002", True, True
            Else
               .SearchMaster "sInsPrmID", pxeDEFAULTINSURANCE, True, True
            End If
         Else
            .SearchMaster "sInsPrmID", pxeDEFAULTINSURANCE, True, True
         End If
      Else
         If .LoadRecord(psTransNox) Then
            txtOthers(0) = IFNull(.Master("xInsurDsc"))
            txtOthers(1) = IFNull(.Master("sOthrInfo"))

            For lnCtr = 0 To .ItemCount - 1
               Select Case lnCtr
               Case 0
                  txtDetail(0) = IFNull(.Detail(lnCtr, "xClientNm"))
                  txtDetail(1) = IFNull(.Detail(lnCtr, "sRelatnDs"))
               Case 1
                  txtDetail(2) = IFNull(.Detail(lnCtr, "xClientNm"))
                  txtDetail(3) = IFNull(.Detail(lnCtr, "sRelatnDs"))
               Case 2
                  txtDetail(4) = IFNull(.Detail(lnCtr, "xClientNm"))
                  txtDetail(5) = IFNull(.Detail(lnCtr, "sRelatnDs"))
               End Select
            Next
            
            xrFrame1.Enabled = False
            xrFrame2.Enabled = False
         Else
            .SearchMaster "sInsPrmID", pxeDEFAULTINSURANCE, True, True
         End If
      
         Call .loadClient
      End If
               
      .Master("sSerialID") = psSerialID
      .Master("dSalesxxx") = pdTransact
      .Master("sSourceNo") = psTransNox
      .Master("sSourceCd") = pxeSOURCECODE
      
      txtField(0) = Format(Mid(.Master("sTransNox"), 2), "@@@-@@-@@@@@@")
   End With
End Sub

Private Sub InitButton(ByVal Value As Integer)
   cmdButton(0).Visible = Value <> xeModeReady
   cmdButton(2).Visible = Value <> xeModeReady
   
   cmdButton(1).Visible = Value = xeModeUpdate
   cmdButton(5).Visible = Value = xeModeUpdate
End Sub

Private Sub InitFields()
   Dim lotxt As TextBox
   
   For Each lotxt In txtField
      lotxt = ""
   Next
   
   For Each lotxt In txtOthers
      lotxt = ""
   Next
   
   For Each lotxt In txtDetail
      lotxt = ""
   Next

   cmdButton(5).Visible = pbViewOnly
   txtOthers(0).Locked = pbViewOnly
   txtOthers(1).Locked = pbViewOnly
End Sub

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

Private Sub oTrans_DetailRetreive(ByVal Index As Integer, ByVal Value As Variant)
   Select Case Index
   Case 0, 2, 4
      txtDetail(Index) = Value
   Case 80
      txtDetail(1) = Value
      oTrans.Benificiary1 = txtDetail(1)
   Case 81
      txtDetail(3) = Value
      oTrans.Benificiary2 = txtDetail(3)
   Case 82
      txtDetail(5) = Value
      oTrans.Benificiary3 = txtDetail(5)
   End Select
End Sub

Private Sub oTrans_MasterRetreive(ByVal Index As Integer, ByVal Value As Variant)
   Select Case Index
   Case 5
      txtOthers(0) = Value
   Case 9
      txtOthers(1) = Value
   End Select
End Sub

Private Sub oTrans_OthersRetreive(ByVal Index As Integer, ByVal Value As Variant)
   Select Case Index
   Case 0
   Case Else
      txtField(Index) = IFNull(Value)
   End Select
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = p_oAppDriver.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   Select Case Index
   Case 0, 1
   Case 2, 3
      If Trim(oTrans.Detail(0, "sClientID")) = "" Or Trim(oTrans.Detail(0, "sRelatnID")) = "" Then
         MsgBox "Invalid Info for Previous Beneficiary. Verify your Entry.", vbCritical, "Warning"
         txtDetail(0).SetFocus
      End If
   Case 4, 5
      If Trim(oTrans.Detail(1, "sClientID")) = "" Or Trim(oTrans.Detail(1, "sRelatnID")) = "" Then
         MsgBox "Invalid Info for Previous Beneficiary. Verify your Entry.", vbCritical, "Warning"
         
         If txtDetail(0) = "" Then
            txtDetail(0).SetFocus
         Else
            txtDetail(2).SetFocus
         End If
      End If
   End Select
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3, vbKeyReturn
      Select Case Index
      Case 0
         oTrans.Detail(1, Index) = txtDetail(Index)
      Case 2
         oTrans.Detail(2, Index) = txtDetail(Index)
      Case 4
         oTrans.Detail(3, Index) = txtDetail(Index)
      Case 1
         oTrans.Detail(0, "sRelatnID") = txtDetail(Index)
      Case 3
         oTrans.Detail(1, "sRelatnID") = txtDetail(Index)
      Case 5
         oTrans.Detail(2, "sRelatnID") = txtDetail(Index)
      End Select
   End Select
      
   KeyCode = 0
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
   txtDetail(Index).BackColor = p_oAppDriver.getColor("EB")
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = p_oAppDriver.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If Index = 1 Then
         Call oTrans.getClient(txtField(Index))
         Call txtField_LostFocus(1)
      End If
   End Select
   KeyCode = 0
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtField(Index).BackColor = p_oAppDriver.getColor("EB")
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = p_oAppDriver.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 0 Then
      Select Case KeyCode
      Case vbKeyReturn
         oTrans.SearchMaster 5, txtOthers(Index), False, False
      Case vbKeyF3
         oTrans.SearchMaster 5, txtOthers(Index), True, False
      End Select
   End If
   KeyCode = 0
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   txtOthers(Index).BackColor = p_oAppDriver.getColor("EB")
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
   Case 1
      oTrans.Master("sOthrInfo") = txtOthers(1)
   End Select
End Sub
