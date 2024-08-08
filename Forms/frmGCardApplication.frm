VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmGCardApplication 
   BorderStyle     =   0  'None
   Caption         =   "GCard Application"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6660
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   11748
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmGCardApplication.frx":0000
         Left            =   8505
         List            =   "frmGCardApplication.frx":0002
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   375
         Width           =   1935
      End
      Begin VB.ListBox lstField 
         Height          =   3210
         Index           =   0
         ItemData        =   "frmGCardApplication.frx":0004
         Left            =   165
         List            =   "frmGCardApplication.frx":0026
         Style           =   1  'Checkbox
         TabIndex        =   30
         Top             =   3285
         Width           =   4515
      End
      Begin VB.ListBox lstField 
         Height          =   3210
         Index           =   1
         ItemData        =   "frmGCardApplication.frx":0167
         Left            =   5190
         List            =   "frmGCardApplication.frx":0198
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   3285
         Width           =   5250
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   8505
         TabIndex        =   27
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   8505
         TabIndex        =   25
         Top             =   2205
         Width           =   1935
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1500
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1140
         Width           =   4860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1500
         MaxLength       =   40
         TabIndex        =   21
         Text            =   "MARLON A. SAYSON"
         Top             =   2505
         Width           =   4860
      End
      Begin VB.OptionButton optField 
         Caption         =   "Plus"
         Height          =   270
         Index           =   1
         Left            =   2625
         TabIndex        =   20
         Tag             =   "wt0;fb0"
         Top             =   2190
         Width           =   1020
      End
      Begin VB.OptionButton optField 
         Caption         =   "Premium"
         Height          =   270
         Index           =   0
         Left            =   1515
         TabIndex        =   19
         Tag             =   "wt0;fb0"
         Top             =   2190
         Width           =   1020
      End
      Begin VB.TextBox txtField 
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
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   2
         Top             =   225
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Classification"
         Height          =   195
         Index           =   5
         Left            =   7035
         TabIndex        =   34
         Top             =   435
         Width           =   1290
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REASON OF PURCHASE"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   32
         Top             =   3015
         Width           =   1860
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOURCE INFO"
         Height          =   195
         Index           =   6
         Left            =   5190
         TabIndex        =   31
         Top             =   3045
         Width           =   1350
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FSEC (White)"
         Height          =   195
         Index           =   1
         Left            =   7035
         TabIndex        =   28
         Top             =   2565
         Width           =   960
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FSEC (Yellow)"
         Height          =   195
         Index           =   2
         Left            =   7035
         TabIndex        =   26
         Top             =   2250
         Width           =   1005
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Availed"
         Height          =   195
         Index           =   3
         Left            =   7035
         TabIndex        =   24
         Top             =   1815
         Width           =   1125
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name to Display"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   2550
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Type"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2235
         Width           =   735
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   8505
         TabIndex        =   17
         Top             =   1455
         Width           =   1935
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   8505
         TabIndex        =   16
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   8505
         TabIndex        =   15
         Top             =   825
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Type"
         Height          =   195
         Index           =   10
         Left            =   7035
         TabIndex        =   14
         Top             =   870
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No."
         Height          =   195
         Index           =   0
         Left            =   7035
         TabIndex        =   13
         Top             =   1185
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame No."
         Height          =   195
         Index           =   1
         Left            =   7035
         TabIndex        =   12
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DR No"
         Height          =   195
         Index           =   2
         Left            =   4710
         TabIndex        =   11
         Top             =   870
         Width           =   495
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   5415
         TabIndex        =   10
         Top             =   825
         Width           =   945
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   1500
         TabIndex        =   9
         Top             =   1455
         Width           =   4860
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Motorcycle"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   8505
         TabIndex        =   8
         Top             =   1770
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name(*)"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   1125
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   5
         Top             =   825
         Width           =   1935
      End
      Begin VB.Label lblFieldNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   870
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Tag             =   "et0;ht2"
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applic No."
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
         Left            =   120
         TabIndex        =   3
         Top             =   285
         Width           =   900
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10995
      TabIndex        =   1
      Top             =   1200
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
      Picture         =   "frmGCardApplication.frx":0307
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   10995
      TabIndex        =   0
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
      Picture         =   "frmGCardApplication.frx":0A81
   End
End
Attribute VB_Name = "frmGCardApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As clsFormSkin

Private p_bIsOkey As Boolean
Private p_bIsLoad As Boolean

Private p_oMaster As Recordset
Private p_oClient As clsNeoClient

Private p_oGClntx As clsNeoClient

Private p_oMCSales As clsMCSales
Private p_oAppDrivr As clsAppDriver
Private p_cDigital As String

Public Property Get Master() As Recordset
   If p_bIsOkey Then
      Set Master = p_oMaster
   Else
      Set Master = Nothing
   End If
End Property

Public Property Set Master(foRS As Recordset)
   Set p_oMaster = foRS
End Property

Public Property Set MCSales(foSales As clsMCSales)
   Set p_oMCSales = foSales
End Property

Public Property Set Client(foClient As clsNeoClient)
   Set p_oClient = foClient
   Set p_oGClntx = xCopy(p_oClient)
End Property

Public Property Get GCard_Client() As clsNeoClient
   Set GCard_Client = p_oGClntx
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

'mac 2019.07.09
Public Function IsDigital() As String
   IsDigital = p_cDigital
End Function

Public Function IsOkey() As Boolean
   IsOkey = p_bIsOkey
End Function


Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      'mac 2019.07.09
      'p_cDigital = Check1.Value 'set digital
      'mac 2020.06.20
      p_cDigital = Combo1.ListIndex + 1
   
      If p_oMaster("sNmOnCard") = "" Then
         MsgBox "Invalid Name to Display detected." & vbCrLf & _
                "Please enter the info and try again!", vbOKOnly + vbInformation, "Validation"
         txtField(1).SetFocus
         Exit Sub
      End If
                  
      If p_oMaster("sClientID") = "" Then
         MsgBox "Invalid Card Owner detected." & vbCrLf & _
                "Please enter the info and try again!", vbOKOnly + vbInformation, "Validation"
         txtField(2).SetFocus
         Exit Sub
      End If
                        
      If p_oGClntx.Master("sLastName") = p_oClient.Master("sLastName") And _
         p_oGClntx.Master("sFrstName") = p_oClient.Master("sFrstName") And _
         p_oGClntx.Master("sMiddName") = p_oClient.Master("sMiddName") And _
         p_oGClntx.Master("sSuffixNm") = p_oClient.Master("sSuffixNm") Then
          
         p_oGClntx.InitClient
      End If
                  
      p_bIsOkey = True
      Me.Hide
   Case 1
      p_bIsOkey = False
      Me.Hide
   End Select
End Sub

Private Sub Form_Activate()
   Dim lnCtr As Integer
   
   If Not p_bIsLoad Then
      'Application No
      txtField(0).Text = p_oMaster("sTransNox")
      'Name on Card
      txtField(1).Text = p_oClient.Master("sFrstName") _
                       + " " + IIf(p_oClient.Master("sMiddName") <> "", Left(p_oClient.Master("sMiddName"), 1) + ".", "") _
                       + " " + p_oClient.Master("sLastName") + " " _
                       + Trim(p_oClient.Master("sSuffixNm"))
      txtField(1).Tag = txtField(1).Text
      p_oMaster("sNmOnCard") = Left(UCase(txtField(1).Text), 35)
      txtField(1).Text = p_oMaster("sNmOnCard")
      
      'Client Name
      txtField(2).Text = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
      txtField(2).Tag = txtField(2).Text
      
      'FSEC Yellow
      txtField(3).Text = p_oMaster("nYellowxx")
      'FSEC White
      txtField(4).Text = p_oMaster("nWhitexxx")
      
      lblField(0).Caption = p_oMCSales.Master("dTransact")
      lblField(1).Caption = p_oMCSales.Master("sDRNoxxxx")
      lblField(2).Caption = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")
      lblField(3).Caption = p_oMCSales.Detail(0, "sModelNme")
      lblField(4).Caption = p_oMCSales.Detail(0, "sEngineNo")
      lblField(5).Caption = p_oMCSales.Detail(0, "sFrameNox")
      'check proper source to display whether new / repo...
      lblField(6).Caption = IIf(p_oMaster("sSourceCD") = "M02910000005", "Motorcycle", "Motorcycle - 2H")
      
      'Card Type
      If p_oMaster("cCardType") = 0 Then
         optField(0).Value = True
      Else
         optField(1).Value = True
      End If
   
      For lnCtr = 1 To Len(p_oMaster("sReasonsx"))
         lstField(0).Selected(lnCtr - 1) = IIf(Mid(p_oMaster("sReasonsx"), lnCtr, 1) = "0", False, True)
      Next
      
      For lnCtr = 1 To Len(p_oMaster("sSrceInfo"))
         lstField(1).Selected(lnCtr - 1) = IIf(Mid(p_oMaster("sSrceInfo"), lnCtr, 1) = "0", False, True)
      Next
      
      'set digital
      'Check1.Value = 0
      'mac 2020.06.20
      '  use combo box as card type
      Combo1.Clear
      'Combo1.AddItem "Smartcard", 0
      Combo1.AddItem "Digital", 0
      Combo1.ListIndex = 0
'      If CDate(p_oAppDrivr.getConfiguration("dNoChipGC", p_oAppDrivr.BranchCode)) <> "1900-01-01" Then
'         If CDate(p_oAppDrivr.getConfiguration("dNoChipGC", p_oAppDrivr.BranchCode)) <= CDate(p_oAppDrivr.ServerDate) Then
'            Combo1.AddItem "Non-chip card", 1
'         End If
'      End If
      'end - mac 2020.06.20
      
      p_bIsLoad = True
   End If
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
   Dim lsProcName As String
   
   lsProcName = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm p_oAppDrivr.MDIMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = p_oAppDrivr
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
endProc:
   Exit Sub
errProc:
'   ShowError lsProcName & "( " & " )", True
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If Index = 2 Then
         Call getGCardClient(txtField(Index), False)
         SetNextFocus
      End If
   End Select
   KeyCode = 0
End Sub

Private Sub lstField_Validate(Index As Integer, Cancel As Boolean)
   Dim lnCtr As Integer
   Dim lsValue As String
      
   For lnCtr = 0 To lstField(Index).ListCount - 1
      lsValue = lsValue + IIf(lstField(Index).Selected(lnCtr), "1", "0")
   Next
      
   If Index = 0 Then
      p_oMaster("sReasonsx") = lsValue
   Else
      p_oMaster("sSrceInfo") = lsValue
   End If
End Sub

Private Sub optField_Validate(Index As Integer, Cancel As Boolean)
   p_oMaster("cCardType") = IIf(optField(0).Value, "0", "1")
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim loClient As clsNeoClient
   Select Case Index
   Case 1
      If Len(txtField(Index)) > 35 Then
         MsgBox "Name on Card will be truncated to 35 characters!!!", vbInformation, "INFO"
         txtField(Index) = Left(txtField(Index), 35)
      End If
      p_oMaster("sNmOnCard") = txtField(Index)
   Case 2
'      Call getGCardClient(txtField(Index), False)
   Case 3
      p_oMaster("nYellowxx") = CInt(txtField(Index))
      txtField(Index) = p_oMaster("nYellowxx")
   Case 4
      p_oMaster("nWhitexxx") = CInt(txtField(Index))
      txtField(Index) = p_oMaster("nWhitexxx")
   End Select
End Sub

Private Function getGCardClient(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getGCardClient"
'   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
   End With
   
   If lsValue <> "" Then
      If Trim(LCase(lsValue)) = Trim(txtField(2).Tag) Then GoTo endProc
   Else
      txtField(2).Text = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
      txtField(2).Tag = txtField(2).Text
      Set p_oGClntx = xCopy(p_oClient)
      GoTo endProc
   End If

   lbExist = loClient.SearchClient(lsValue, False)

   If Not lbExist Then
      lasName = GetSplitedName(lsValue)
      loClient.Master("sLastName") = lasName(0)
      loClient.Master("sFrstName") = lasName(1)
   End If

   If loClient.getClient Then
      Set p_oGClntx = loClient
   Else
      Set p_oGClntx = xCopy(p_oClient)
   End If

   p_oMaster("sClientID") = p_oGClntx.Master("sClientID")
   txtField(2).Text = p_oGClntx.Master("sLastName") + ", " + p_oGClntx.Master("sFrstName") + " " + Trim(p_oGClntx.Master("sSuffixNm")) + IIf(Trim(p_oGClntx.Master("sSuffixNm")) = "", "", " ") + p_oGClntx.Master("sMiddName")
   txtField(2).Tag = txtField(2).Text
   
   getGCardClient = True
   
endProc:
   Exit Function
End Function

Function xCopy(foClient As clsNeoClient) As clsNeoClient
   Dim loClient As clsNeoClient
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_oAppDrivr.BranchCode
      If .InitClient() = False Then Exit Function
      .Master("sLastName") = p_oClient.Master("sLastName")
      .Master("sFrstname") = p_oClient.Master("sFrstname")
      .Master("sMiddName") = p_oClient.Master("sMiddName")
      .Master("sSuffixNm") = p_oClient.Master("sSuffixNm")
      .Master("cGenderCd") = p_oClient.Master("cGenderCd")
      .Master("cCvilStat") = p_oClient.Master("cCvilStat")
      .Master("dBirthDte") = p_oClient.Master("dBirthDte")
      .Master("sBirthPlc") = p_oClient.Master("sBirthPlc")
      .Master("sHouseNox") = p_oClient.Master("sHouseNox")
      .Master("sAddressx") = p_oClient.Master("sAddressx")
      .Master("sTownIDxx") = p_oClient.Master("sTownIDxx")
      .Master("sBrgyIDxx") = p_oClient.Master("sBrgyIDxx")
      .Master("sPhoneNox") = p_oClient.Master("sPhoneNox")
      .Master("sMobileNo") = p_oClient.Master("sMobileNo")
      .Master("sEmailAdd") = p_oClient.Master("sEmailAdd")
   End With
   
   Set xCopy = loClient
End Function

