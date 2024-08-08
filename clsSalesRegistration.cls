VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSalesRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Note: Replacement transaction is currently not incorported with the current logic
' Transaction Type:
'  0 = "Gift"           X
'  1 = "Sales"          X
'  2 = "Replacement"    X  xxx
'
' Payment Form
'  0 = "Cash"
'  1 = "Cash Balance"
'  2 = "Installment"
'  3 = "Term"
'  4 = "Term with Registration"


Option Explicit

Private Const pxeMODULENAME = "clsSalesRegistration"

Private p_oAppDrivr As clsAppDriver

Private p_dTransact As Date
Private p_sReferNox As String
Private p_sClientID As String
Private p_sBranchCd As String
Private p_nEditMode As xeEditMode
Private p_cPaymForm As String
Private p_cTranType As String

Private p_oDetail As New ggcMCRegistration.clsSalesDetail
Private p_aDetail As Collection

Private p_sCoCltID1 As String
Private p_sCoCltID2 As String

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Let PaymForm(fValue As String)
   p_cPaymForm = fValue
End Property

Property Let TranType(fValue As String)
   p_cTranType = fValue
End Property

Property Let Detail(Row As Integer, Index As Variant, Value As Variant)
   Dim lsOldProc As String
   
   lsOldProc = "(Let)Detail(" & Row & ", " & Index & ", " & Value & ")"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   If p_aDetail.Count + 1 < Row Then
      Exit Property
   ElseIf p_aDetail.Count + 1 = Row Then
      With p_aDetail
         Set p_oDetail = New clsSalesDetail
         .Add p_oDetail
         .Item(.Count).SerialID = ""
         .Item(.Count).Insurance = ""
         .Item(.Count).Registration = ""
         .Item(.Count).Amount = 0#
         .Item(.Count).IsNew = ""
         .Item(.Count).OriginalSerial = ""
         .Item(.Count).ReplacedSerial = ""
         .Item(.Count).ReplaceInsurance = 0#
      End With
   End If
      
   Select Case Index
      Case 0, "sserialid"
         p_aDetail(Row).SerialID = Value
      Case 1, "sinstypid"
         p_aDetail(Row).Insurance = Value
      Case 2, "cregistyp"
         p_aDetail(Row).Registration = IIf(IsNull(Value), "x", Value)
      Case 3, "ninsamtxx"
         p_aDetail(Row).Amount = Value
      Case 4, "cmotornew"
         p_aDetail(Row).IsNew = Value
      Case 5, "sprevserl"
         p_aDetail(Row).OriginalSerial = Value
      Case 6, "sreplmcid"
         p_aDetail(Row).ReplacedSerial = Value
      Case 7, "nrepliamt"
         p_aDetail(Row).ReplaceInsurance = Value
   
   End Select
   
endProc:
   Exit Property
errProc:
   ShowError lsOldProc
End Property

Property Get Detail(Row As Integer, Index As Variant) As Variant
   Dim lsOldProc As String
   
   lsOldProc = "(Get)Detail(" & Row & ", " & Index & ")"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   If p_aDetail.Count < Row Then
      Exit Property
   End If
   
   Select Case Index
      Case 0, "sserialid"
         Detail = p_aDetail(Row).SerialID
      Case 1, "sinstypid"
         Detail = p_aDetail(Row).Insurance
      Case 2, "cregistyp"
         Detail = p_aDetail(Row).Registration
      Case 3, "ninsamtxx"
         Detail = p_aDetail(Row).Amount
      Case 4, "cmotornew"
         Detail = p_aDetail(Row).IsNew
      Case 5, "sprevserl"
         Detail = p_aDetail(Row).OriginalSerial
      Case 6, "sreplmcid"
         Detail = p_aDetail(Row).ReplacedSerial
      Case 7, "nrepliamt"
         Detail = p_aDetail(Row).ReplaceInsurance
   End Select

endProc:
   Exit Property
errProc:
   ShowError lsOldProc
End Property

Function InitTransaction(Optional oAppDriver As Variant) As Boolean
   Dim loRS As ADODB.Recordset
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "InitTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   If IsMissing(oAppDriver) Then
      If isAppDriverOK(p_oAppDrivr) = False Then
         MsgBox pxeMODULENAME & " » Invalid driver passed detected!!!" & vbCrLf & vbCrLf & _
                "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   Else
      If isAppDriverOK(p_oAppDrivr) = False Then
         MsgBox pxeMODULENAME & " » Wrong driver type detected!!!" & vbCrLf & vbCrLf & _
                "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      Set p_oAppDrivr = oAppDriver
   End If
   
   InitTransaction = InitDetail

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function InitDetail() As Boolean
   Set p_aDetail = New Collection
   If p_aDetail.Count = 0 Then
      Set p_oDetail = New clsSalesDetail
      p_aDetail.Add p_oDetail
   End If
   
   Do Until p_aDetail.Count = 1
      p_aDetail.Remove 1
   Loop
      
   p_aDetail(1).SerialID = ""
   p_aDetail(1).Insurance = ""
   p_aDetail(1).Registration = ""
   p_aDetail(1).Amount = 0#
   p_aDetail(1).IsNew = ""
   p_aDetail(1).OriginalSerial = ""
   p_aDetail(1).ReplacedSerial = ""
   
   InitDetail = True
End Function

Function SaveTransaction( _
         ByVal fsReferNox As String, _
         ByVal fdTransact As Date, _
         ByVal fsClientID As String, _
         ByVal fnEditMode As xeEditMode _
         ) As Boolean
   Dim loObj As clsSalesDetail
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim loRS As ADODB.Recordset
   Dim lsNotIn As String
   Dim lasClient() As String
  
   
   lsOldProc = "SaveTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   p_sReferNox = fsReferNox
   p_dTransact = fdTransact
   
   lasClient = Split(fsClientID, "»")
   
   If UBound(lasClient) = 2 Then
      p_sCoCltID2 = lasClient(2)
   Else
      p_sCoCltID2 = ""
   End If
   
   If UBound(lasClient) >= 1 Then
      p_sCoCltID1 = lasClient(1)
   Else
      p_sCoCltID1 = ""
   End If
   
   p_sClientID = lasClient(0)
   
   p_nEditMode = fnEditMode
   p_sBranchCd = Left(p_sReferNox, Len(p_oAppDrivr.BranchCode))
   
   lsNotIn = ""
   
   For Each loObj In p_aDetail
      If p_nEditMode = xeModeAddNew Then
         If Trim(loObj.SerialID) <> "" Then
            'A Replacement
            If p_cTranType = "2" Then
               If SaveReplacement(loObj) = False Then GoTo endProc
            ' Regular Sales with Insurance paid
            ' Regular Sales with Term w/ registration (no insurance paid)
            ElseIf loObj.Amount <> 0 Then
               If Trim(loObj.Registration) <> "" Then
                  If SaveInsurance(loObj.SerialID, loObj.Registration & "»" & loObj.IsNew, loObj.Insurance) = False Then GoTo endProc
                  If SaveRegistration(loObj.SerialID, loObj.Registration, loObj.IsNew, loObj.Amount) = False Then GoTo endProc
                  SaveLTOPool loObj, "SRg", loObj.Amount
                  If loObj.IsNew <> 1 Then
                     If SaveDeedSale(loObj.SerialID, 0) = False Then GoTo endProc
                  End If
               Else
                  If SaveDeedSale(loObj.SerialID, loObj.Amount) = False Then GoTo endProc
                  SaveLTOPool loObj, "SDS", loObj.Amount
               End If
            End If
         End If
      ElseIf p_nEditMode = xeModeUpdate Then
         If Trim(loObj.SerialID) <> "" Then
            '  Replacement
            If p_cTranType = "2" Then
               If SaveReplacement(loObj) = False Then GoTo endProc
            ElseIf loObj.Amount <> 0 Then
               If Trim(loObj.Registration) <> "" Then
                  'Well it just an update...
                  If loObj.IsNew <> 1 Then
                     If SaveDeedSale(loObj.SerialID, 0, loObj.OriginalSerial) = False Then GoTo endProc
                  Else
                     Call UnSaveDeedSale(loObj.OriginalSerial)
                  End If
                  If SaveInsurance(loObj.SerialID, loObj.Registration & "»" & loObj.IsNew, loObj.Insurance, loObj.OriginalSerial) = False Then GoTo endProc
                  If SaveRegistration(loObj.SerialID, loObj.Registration, loObj.IsNew, loObj.Amount, loObj.OriginalSerial) = False Then GoTo endProc
                  SaveLTOPool loObj, "SRg", loObj.Amount
               Else
                  'The record might have insurance and registration at some time
                  'so try to delete them
                  If UnSaveInsurance(loObj.OriginalSerial) = False Then GoTo endProc
                  If UnSaveRegistration(loObj.OriginalSerial) = False Then GoTo endProc
                  If SaveDeedSale(loObj.SerialID, loObj.Amount) = False Then GoTo endProc
                  DeleteLTOPool loObj.SerialID, "SRg"
                  DeleteLTOPool loObj.SerialID, "SDS"
               End If
            Else
               'If it has no insurance amount and trantype is not replacement then delete them since at one time
               'it might have insurance and registration/deed of sale
               If UnSaveInsurance(loObj.SerialID) = False Then GoTo endProc
               If UnSaveRegistration(loObj.SerialID) = False Then GoTo endProc
               If UnSaveDeedSale(loObj.SerialID) = False Then GoTo endProc
               DeleteLTOPool loObj.SerialID, "SRg"
            End If
         End If
      ElseIf p_nEditMode = xeModeDelete Then
         If p_cTranType = "2" Then
            If UnSaveReplacement(loObj) = False Then GoTo endProc
         Else
            If UnSaveInsurance(loObj.SerialID) = False Then GoTo endProc
            If UnSaveRegistration(loObj.SerialID) = False Then GoTo endProc
            If UnSaveDeedSale(loObj.SerialID) = False Then GoTo endProc
            DeleteLTOPool loObj.SerialID, "SRg"
         End If
      End If
      
      'Record Process serial here
      lsNotIn = lsNotIn & ", " & strParm(Trim(loObj.SerialID))
   Next
   
   'Delete Record that are not included in the list...
   If p_nEditMode = xeModeUpdate Or p_nEditMode = xeModeDelete Then
      'Delete Records in Registration and Insurance
      lsSQL = " SELECT sSerialID" & _
              " FROM MC_Registration" & _
              " WHERE sReferNox = " & strParm(p_sReferNox) & _
                " AND sSourceCD = " & strParm("SRg") & _
                " AND sSerialID NOT IN (" & Mid(lsNotIn, 2) & ")"
      
      Set loRS = New ADODB.Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      Do Until loRS.EOF
         If UnSaveInsurance(loRS("sSerialID"), True) = False Then GoTo endProc
         If UnSaveRegistration(loRS("sSerialID"), True) = False Then GoTo endProc
         DeleteLTOPool loRS("sSerialID"), "SRg"
         loRS.MoveNext
      Loop

      'Delete Records in DeedOfSale
      lsSQL = " SELECT sSerialID" & _
              " FROM MC_DeedSale_Expense" & _
              " WHERE sReferNox = " & strParm(p_sReferNox) & _
                " AND sSourceCD = " & strParm("SDS") & _
                " AND sSerialID NOT IN (" & Mid(lsNotIn, 2) & ")"
      Set loRS = New ADODB.Recordset
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      Do Until loRS.EOF
         If UnSaveDeedSale(loRS("sSerialID"), True) = False Then GoTo endProc
         DeleteLTOPool loRS("sSerialID"), "SDS"
         loRS.MoveNext
      Loop
   End If
   
   SaveTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function SaveRegistration( _
                 ByVal fsSerialID As String, _
                 ByVal fcRegisTyp As String, _
                 ByVal fcMotorNew As String, _
                 ByVal fnTranAmtx As Currency, _
                 Optional ByVal fsPrevSerl As String = "" _
                 ) As Boolean
   
   Dim lsOldProc As String
   Dim loLTOPool As clsMCLTOPool
   Dim loMCRegs As clsMCRegistration
   Dim loMCRegEx As clsMCRegExpense
   
   lsOldProc = "SaveRegistration"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
      
   If fsPrevSerl = "" Then fsPrevSerl = fsSerialID
   
   'Save Registration
   Set loMCRegs = New clsMCRegistration
   Set loMCRegs.AppDriver = p_oAppDrivr
   loMCRegs.Branch = p_sBranchCd
   If loMCRegs.InitRecord = False Then Exit Function
'130      If p_nEditMode = xeModeAddNew Then
'140         If loMCRegs.NewRecord = False Then Exit Function
'150      Else
      If loMCRegs.SearchRecord(p_sReferNox, "SRg", fsPrevSerl) = False Then
         If loMCRegs.NewRecord = False Then Exit Function
      Else
         If loMCRegs.UpdateRecord = False Then Exit Function
      End If
'210      End If
   loMCRegs.Master("dPurchase") = p_dTransact
   loMCRegs.Master("sReferNox") = p_sReferNox
   loMCRegs.Master("sSourceCd") = "SRg"
   loMCRegs.Master("sSerialID") = fsSerialID
   loMCRegs.Master("sClientID") = p_sClientID
   loMCRegs.Master("sCoCltID1") = p_sCoCltID1
   loMCRegs.Master("sCoCltID2") = p_sCoCltID2
   loMCRegs.Master("cRegisTyp") = fcRegisTyp
   loMCRegs.Master("sORNoxxxx") = p_cPaymForm
'   loMCRegs.Master("cRegisFrm") = IIf(fcMotorNew = "1", "0", "1")
   loMCRegs.Master("nRegisAmt") = fnTranAmtx
   If loMCRegs.SaveRecord = False Then Exit Function
   
   'Save Registration Expense
   Set loMCRegEx = New clsMCRegExpense
   Set loMCRegEx.AppDriver = p_oAppDrivr
   loMCRegEx.Branch = p_sBranchCd
   If loMCRegEx.InitRecord = False Then Exit Function
'360      If p_nEditMode = xeModeAddNew Then
'370         If loMCRegEx.NewRecord = False Then Exit Function
'380      Else
      If loMCRegEx.SearchRecord(loMCRegs.Master("sTransNox")) = False Then
         If loMCRegEx.NewRecord = False Then Exit Function
      Else
         If loMCRegEx.UpdateRecord = False Then Exit Function
      End If
'440      End If
   loMCRegEx.Master("sReferNox") = loMCRegs.Master("sTransNox")
   loMCRegEx.Master("cMotorNew") = fcMotorNew
   If loMCRegEx.SaveRecord = False Then Exit Function
          
   SaveRegistration = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function SaveInsurance( _
                 ByVal fsSerialID As String, _
                 ByVal fcRegisTyp As String, _
                 ByVal fsInsTypID As String, _
                 Optional ByVal fsPrevSerl As String = "" _
                 ) As Boolean
   Dim loInsur As clsMCInsurance
   Dim lcRegisFrm As String
   Dim facRegisTyp() As String
   Dim lsOldProc As String
   
   lsOldProc = "SaveInsurance()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   facRegisTyp = Split(fcRegisTyp, "»")
   If fsPrevSerl = "" Then fsPrevSerl = fsSerialID
         
   fcRegisTyp = facRegisTyp(0)
   lcRegisFrm = facRegisTyp(1)
   
   Set loInsur = New clsMCInsurance
   Set loInsur.AppDriver = p_oAppDrivr
   loInsur.Branch = p_sBranchCd
   If loInsur.InitRecord = False Then Exit Function
'160      If p_nEditMode = xeModeAddNew Then
'170         If loInsur.NewRecord = False Then Exit Function
'180      Else
      If loInsur.SearchRecord(p_sReferNox, "Sl", fsPrevSerl) = False Then
         If loInsur.NewRecord = False Then Exit Function
         loInsur.Master("cTranStat") = xeStateOpen
         loInsur.Master("cRecvCOCx") = xeNo
      Else
         If loInsur.UpdateRecord = False Then Exit Function
      End If
'260      End If
   loInsur.Master("sReferNox") = p_sReferNox
   loInsur.Master("sSourceCD") = "Sl"
   loInsur.Master("sClientID") = p_sClientID
   loInsur.Master("sCoCltID1") = p_sCoCltID1
   loInsur.Master("sCoCltID2") = p_sCoCltID2
   loInsur.Master("sSerialID") = fsSerialID
   loInsur.Master("dTransact") = p_dTransact
   loInsur.Master("sInsTypID") = fsInsTypID
   loInsur.Master("cRegisTyp") = fcRegisTyp
   loInsur.Master("cRegisFrm") = lcRegisFrm
   If loInsur.SaveRecord = False Then Exit Function
   
   SaveInsurance = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function SaveDeedSale( _
                 fsSerialID As String, _
                 ByVal fnTranAmtx As Currency, _
                 Optional ByVal fsPrevSerl As String = "" _
                 ) As Boolean
   Dim lsOldProc As String
   Dim loLTOPool As clsMCLTOPool
   Dim loDeeds As clsMCDeedSaleExpense
      
   lsOldProc = "SaveDeedSale"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   If fsPrevSerl = "" Then fsPrevSerl = fsSerialID
   
   Set loDeeds = New clsMCDeedSaleExpense
   Set loDeeds.AppDriver = p_oAppDrivr
   loDeeds.Branch = p_sBranchCd
   If loDeeds.InitRecord = False Then Exit Function
   If p_nEditMode = xeModeAddNew Then
      If loDeeds.NewRecord = False Then Exit Function
   Else
      If loDeeds.SearchRecord(p_sReferNox, "SDS", fsPrevSerl) = False Then
         If loDeeds.NewRecord = False Then Exit Function
      Else
         If loDeeds.UpdateRecord = False Then Exit Function
      End If
   End If
   loDeeds.Master("dTransact") = p_dTransact
   loDeeds.Master("sReferNox") = p_sReferNox
   loDeeds.Master("sSourceCd") = "SDS"
   loDeeds.Master("sClientID") = p_sClientID
   loDeeds.Master("sCoCltID1") = p_sCoCltID1
   loDeeds.Master("sCoCltID2") = p_sCoCltID2
   loDeeds.Master("sSerialID") = fsSerialID
   loDeeds.Master("cDeedSPrc") = xeNo
   loDeeds.Master("nAmtPaidx") = fnTranAmtx
   If loDeeds.SaveRecord = False Then Exit Function
   
   SaveDeedSale = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function UnSaveRegistration( _
                 ByVal fsSerialID As String, _
                 Optional ByVal fbIsForced As Boolean = False _
                 ) As Boolean
   Dim lsOldProc As String
   Dim loLTOPool As clsMCLTOPool
   Dim loMCRegs As clsMCRegistration
   Dim loMCRegEx As clsMCRegExpense
      
   lsOldProc = "UnSaveRegistration()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   'Load MC Registration
   Set loMCRegs = New clsMCRegistration
   Set loMCRegs.AppDriver = p_oAppDrivr
   loMCRegs.Branch = p_sBranchCd
   If loMCRegs.InitRecord = False Then Exit Function
   If loMCRegs.SearchRecord(p_sReferNox, _
                            "SRg", _
                            fsSerialID) = False Then
      UnSaveRegistration = Not fbIsForced
      GoTo endProc
   End If
   
   'Delete Registration Expense
   Set loMCRegEx = New clsMCRegExpense
   Set loMCRegEx.AppDriver = p_oAppDrivr
   loMCRegEx.Branch = p_sBranchCd
   If loMCRegEx.InitRecord = False Then Exit Function
   If loMCRegEx.SearchRecord(loMCRegs.Master("sTransNox")) = False Then
      UnSaveRegistration = Not fbIsForced
      GoTo endProc
   End If
   If loMCRegEx.DeleteRecord = False Then Exit Function
   
   'Delete Registration
   If loMCRegs.DeleteRecord = False Then Exit Function
   
   UnSaveRegistration = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function UnSaveInsurance( _
                 ByVal fsSerialID As String, _
                 Optional ByVal fbIsForced As Boolean = False _
                 ) As Boolean
   Dim lsOldProc As String
   Dim loInsure As clsMCInsurance
   
   lsOldProc = "UnSaveInsurance"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   UnSaveInsurance = False
   
   'Delete Deed Sale Expense
   Set loInsure = New clsMCInsurance
   Set loInsure.AppDriver = p_oAppDrivr
   loInsure.Branch = p_sBranchCd
   If loInsure.InitRecord = False Then Exit Function
   If loInsure.SearchRecord(p_sReferNox, _
                            "Sl", _
                            fsSerialID) = False Then
      UnSaveInsurance = Not fbIsForced
      GoTo endProc
   End If
   If loInsure.DeleteRecord = False Then Exit Function
   
   UnSaveInsurance = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function UnSaveDeedSale( _
                 ByVal fsSerialID As String, _
                 Optional ByVal fbIsForced As Boolean = False _
                 ) As Boolean
   Dim lsOldProc As String
   Dim loLTOPool As clsMCLTOPool
   Dim loDeeds As clsMCDeedSaleExpense
   Dim lnAmtPaidx As Currency
   
   lsOldProc = "UnSaveDeedSale()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   'Delete Deed Sale Expense
   Set loDeeds = New clsMCDeedSaleExpense
   Set loDeeds.AppDriver = p_oAppDrivr
   loDeeds.Branch = p_sBranchCd
   If loDeeds.InitRecord = False Then Exit Function
   If loDeeds.SearchRecord(p_sReferNox, _
                           "SDS", _
                           fsSerialID) = False Then
      UnSaveDeedSale = Not fbIsForced
      GoTo endProc
   Else
      lnAmtPaidx = loDeeds.Master("nAmtPaidx")
      If loDeeds.DeleteRecord = False Then GoTo endProc
   End If
   
   UnSaveDeedSale = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      With Err
         .Raise .Number, .Source, .Description
      End With
   End With
End Sub

Private Sub SaveLTOPool(ByVal foObj As Object, ByVal fsSourceCd As String, ByVal fnTranAmtx As Currency)
   Dim loLTOPool As clsMCLTOPool
   
   Set loLTOPool = New clsMCLTOPool
   Set loLTOPool.AppDriver = p_oAppDrivr
   loLTOPool.TranDate = p_dTransact
   loLTOPool.SerialID = foObj.SerialID
   loLTOPool.PrevSerial = foObj.OriginalSerial
   Call loLTOPool.SaveCollection( _
        p_sBranchCd, _
        p_sReferNox, _
        fsSourceCd, _
        fnTranAmtx _
        )
   Set loLTOPool = Nothing

End Sub

Private Sub DeleteLTOPool(ByVal sSerialID As String, ByVal sSourceCd As String)
   Dim loLTOPool As clsMCLTOPool
   
   'Update LTO Pool
   Set loLTOPool = New clsMCLTOPool
   Set loLTOPool.AppDriver = p_oAppDrivr
   loLTOPool.SerialID = sSerialID
   loLTOPool.TranDate = p_dTransact
   Call loLTOPool.DeleteCollection( _
        p_sBranchCd, _
        p_sReferNox, _
        sSourceCd _
        )
   Set loLTOPool = Nothing
End Sub

Private Function SaveReplacement(foObj As Object) As Boolean
   Dim lsSQL As String
   Dim loRSSales As ADODB.Recordset
   Dim loRSRegis As ADODB.Recordset
   Dim loRSRegisx As ADODB.Recordset
   Dim lsOldProc As String
   Dim lnTranStat As xeTransactionStatus
   Dim lbHasRegis As Boolean

   lsOldProc = "SaveReplacement()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   'Get Previous Sales History
   lsSQL = "SELECT a.sTransNox" & _
          " FROM MC_SO_Master a, MC_SO_Detail b" & _
          " WHERE a.sTransNox = b.sTransNox" & _
            " AND a.sTransNox <> " & strParm(p_sReferNox) & _
            " AND a.sClientID = " & strParm(p_sClientID) & _
            " AND b.sSerialID = " & strParm(foObj.ReplacedSerial)
            Debug.Print lsSQL
   Set loRSSales = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   'Determine what should be the status
   lnTranStat = IIf(foObj.Amount = 0, xeStateCancelled, xeStateUnknown)
      
   'Get Previous Registration
   lsSQL = "SELECT b.sTransNox, b.cInsurPrc, b.cTranStat" & _
          " FROM MC_Registration a, MC_Registration_Expense b" & _
          " WHERE a.sSourceCd = 'SRg'" & _
            " AND a.sTransNox = b.sReferNox" & _
            " AND a.sReferNox = " & strParm(loRSSales("sTransNox")) & _
            " AND a.sSerialID = " & strParm(foObj.ReplacedSerial) & _
          " ORDER BY a.sTransNox DESC"
   Set loRSRegisx = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
   'Don't continue the processing if registration/insurance processing is on going
   If Not loRSRegisx.EOF Then
      If loRSRegisx("cInsurPrc") = xeYes And foObj.Amount = 0 Then
         MsgBox "Registration is Being Processed!!!" & vbCrLf & _
             "Customer is required to pay the insurance and registration amount!!" & vbCrLf & vbCrLf & _
             "You can inform GGC SEG/SSG for detail!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      'Cancel Previous Registration
      If lnTranStat <> loRSRegisx("cTranStat") Then
         'This command is faster than loading the entire class
         lsSQL = "UPDATE MC_Registration_Expense SET" & _
                     " cTranStat = " & strParm(lnTranStat) & _
                " WHERE sTransNox = " & strParm(loRSRegisx("sTransNox"))
         p_oAppDrivr.Execute lsSQL, "MC_Registration_Expense", p_sBranchCd
      End If
   End If
             
   'Get Previous Insurance
   lsSQL = "SELECT sInsurNox, cTranstat, sInsTypID, cRegisTyp" & _
          " FROM MC_Insurance" & _
          " WHERE sSourceCd = 'Sl'" & _
            " AND sReferNox = " & strParm(loRSSales("sTransNox")) & _
            " AND sSerialID = " & strParm(foObj.ReplacedSerial) & _
          " ORDER BY sInsurNox DESC"
   Set loRSRegis = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
   If Not loRSRegis.EOF Then
      If lnTranStat <> loRSRegis("cTranStat") Then
         lsSQL = "UPDATE MC_Insurance SET" & _
                     " cTranStat = " & strParm(lnTranStat) & _
                " WHERE sInsurNox = " & strParm(loRSRegis("sInsurNox"))
         p_oAppDrivr.Execute lsSQL, "MC_Insurance", p_sBranchCd
      End If
      foObj.Registration = loRSRegis("cRegisTyp")
      foObj.Insurance = loRSRegis("sInsTypID")
   End If
      
   'Get Previous Deed Sale Expense
   lsSQL = "SELECT sTransNox, cTranStat" & _
          " FROM MC_DeedSale_Expense" & _
          " WHERE sSourceCd = 'SDS'" & _
            " AND sReferNox = " & strParm(loRSSales("sTransNox")) & _
            " AND sSerialID = " & strParm(foObj.ReplacedSerial) & _
          " ORDER BY sTransNox DESC"
   Set loRSRegis = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
   If Not loRSRegis.EOF Then
      'Cancel Previous Registration
      If lnTranStat <> loRSRegis("cTranStat") Then
         'This command is faster than loading the entire class
         lsSQL = "UPDATE MC_DeedSale_Expense SET" & _
                     " cTranStat = " & strParm(lnTranStat) & _
                " WHERE sTransNox = " & strParm(loRSRegis("sTransNox"))
         p_oAppDrivr.Execute lsSQL, "MC_DeedSale_Expense", p_sBranchCd
      End If
   End If
            
   If loRSRegisx.EOF Then
      lbHasRegis = True
   Else
      lbHasRegis = loRSRegisx("cInsurPrc") = xeNo Or _
                        loRSRegisx("cInsurPrc") = "" Or _
                        (loRSRegisx("cInsurPrc") = xeYes And (foObj.Amount = 0 + foObj.ReplaceInsurance) > 0)
   End If
            
   ' Note: Make sure that user will not fill the insurance amount of mcsales detail
   '       if the customer does not pay for new registration of the replaced mc.
   If p_nEditMode = xeModeAddNew Then
'430         If foObj.Amount = 0 + foObj.ReplaceInsurance > 0 Then
      If lbHasRegis Then
         If Trim(foObj.Registration) <> "" Then
            If SaveInsurance( _
               foObj.SerialID, _
               foObj.Registration & "»" & foObj.IsNew, _
               foObj.Insurance _
               ) = False Then GoTo endProc
            
            If SaveRegistration( _
               foObj.SerialID, _
               foObj.Registration, _
               foObj.IsNew, _
               IIf(foObj.Amount = 0, foObj.ReplaceInsurance, foObj.Amount) _
               ) = False Then GoTo endProc
            
            SaveLTOPool foObj, "SRg", foObj.Amount
            
            If foObj.IsNew <> 1 Then
               If SaveDeedSale(foObj.SerialID, 0) = False Then GoTo endProc
            End If
         Else
            If SaveDeedSale(foObj.SerialID, IIf(foObj.Amount = 0, foObj.ReplaceInsurance, IIf(foObj.Amount = 0, foObj.ReplaceInsurance, foObj.Amount))) = False Then GoTo endProc
            SaveLTOPool foObj, "SDS", foObj.Amount
         End If
      End If
   ElseIf p_nEditMode = xeModeUpdate Then
'570         If foObj.Amount = 0 + foObj.ReplaceInsurance > 0 Then
      If lbHasRegis Then
         If Trim(foObj.Registration) <> "" Then
            'Well it just an update...
            If foObj.IsNew <> 1 Then
               If SaveDeedSale( _
                  foObj.SerialID, _
                  0, _
                  foObj.OriginalSerial _
                  ) = False Then GoTo endProc
            Else
               Call UnSaveDeedSale(foObj.OriginalSerial)
            End If
            If SaveInsurance( _
               foObj.SerialID, _
               foObj.Registration & "»" & foObj.IsNew, _
               foObj.Insurance, _
               foObj.OriginalSerial _
               ) = False Then GoTo endProc
            
            If SaveRegistration( _
               foObj.SerialID, _
               foObj.Registration, _
               foObj.IsNew, _
               IIf(foObj.Amount = 0, foObj.ReplaceInsurance, foObj.Amount), _
               foObj.OriginalSerial _
               ) = False Then GoTo endProc
            SaveLTOPool foObj, "SRg", foObj.Amount
         End If
      Else
         'The record might have insurance and registration at some time
         'so try to delete them
         If UnSaveInsurance(foObj.OriginalSerial) = False Then GoTo endProc
         If UnSaveRegistration(foObj.OriginalSerial) = False Then GoTo endProc
         If SaveDeedSale(foObj.SerialID, IIf(foObj.Amount = 0, foObj.ReplaceInsurance, foObj.Amount)) = False Then GoTo endProc
         DeleteLTOPool foObj.SerialID, "SRg"
         DeleteLTOPool foObj.SerialID, "SDS"
      End If
   End If
   
   SaveReplacement = True
         
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function UnSaveReplacement(foObj As Object) As Boolean
   Dim lsSQL As String
   Dim loRSSales As ADODB.Recordset
   Dim loRSRegis As ADODB.Recordset
   Dim lsOldProc As String
   Dim lnTranStat As xeTransactionStatus

   lsOldProc = "UnSaveReplacement()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   'Get Previous Sales History
   lsSQL = "SELECT a.sTransNox" & _
          " FROM MC_SO_Master a, MC_SO_Detail b" & _
          " WHERE a.sTransNox = b.sTransNox" & _
            " AND a.sTransNox <> " & strParm(p_sReferNox) & _
            " AND a.sClientID = " & strParm(p_sClientID) & _
            " AND b.sSerialID = " & strParm(foObj.ReplacedSerial)
   Set loRSSales = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   'Determine what should be the status
   lnTranStat = IIf(foObj.Amount = 0, xeStateOpen, xeStatePosted)
      
   'Get Previous Registration
   lsSQL = "SELECT b.sTransNox, b.cInsurPrc, b.cTranStat" & _
          " FROM MC_Registration a, MC_Registration_Expense b" & _
          " WHERE a.sSourceCd = 'SRg'" & _
            " AND a.sTransNox = b.sReferNox" & _
            " AND a.sReferNox = " & strParm(loRSSales("sTransNox")) & _
            " AND a.sSerialID = " & strParm(foObj.ReplacedSerial)
   Set loRSRegis = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
   If Not loRSRegis.EOF Then
      'Cancel Previous Registration
      If lnTranStat <> loRSRegis("cTranStat") Then
         'This command is faster than loading the entire class
         lsSQL = "UPDATE MC_Registration_Expense SET" & _
                     " cTranStat = " & strParm(lnTranStat) & _
                " WHERE sTransNox = " & strParm(loRSRegis("sTransNox"))
         p_oAppDrivr.Execute lsSQL, "MC_Registration_Expense", p_sBranchCd
      End If
   End If
             
   'Get Previous Insurance
   lsSQL = "SELECT sInsurNox, cTranstat" & _
          " FROM MC_Insurance" & _
          " WHERE sSourceCd = 'Sl'" & _
            " AND sReferNox = " & strParm(loRSSales("sTransNox")) & _
            " AND sSerialID = " & strParm(foObj.ReplacedSerial)
   Set loRSRegis = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
   If Not loRSRegis.EOF Then
      If lnTranStat <> loRSRegis("cTranStat") Then
         lsSQL = "UPDATE MC_Insurance SET" & _
                     " cTranStat = " & strParm(lnTranStat) & _
                " WHERE sInsurNox = " & strParm(loRSRegis("sInsurNox"))
         p_oAppDrivr.Execute lsSQL, "MC_Insurance", p_sBranchCd
      End If
   End If
      
   'Get Previous Deed Sale Expense
   lsSQL = "SELECT sTransNox, cTranStat" & _
          " FROM MC_DeedSale_Expense" & _
          " WHERE sSourceCd = 'SDS'" & _
            " AND sReferNox = " & strParm(loRSSales("sTransNox")) & _
            " AND sSerialID = " & strParm(foObj.ReplacedSerial)
   Set loRSRegis = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      
   If Not loRSRegis.EOF Then
      'Cancel Previous Registration
      If lnTranStat <> loRSRegis("cTranStat") Then
         'This command is faster than loading the entire class
         lsSQL = "UPDATE MC_DeedSale_Expense SET" & _
                     " cTranStat = " & strParm(lnTranStat) & _
                " WHERE sTransNox = " & strParm(loRSRegis("sTransNox"))
         p_oAppDrivr.Execute lsSQL, "MC_DeedSale_Expense", p_sBranchCd
      End If
   End If

   UnSaveReplacement = True
         
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function
