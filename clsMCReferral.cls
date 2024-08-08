VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCReferral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     MC Referral
'        - Use in entry and maintenance of MC Referral Monitoring...
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
Option Explicit

Private Const pxeMODULENAME = "clsMCReferral"

Private p_oAppDrivr As clsAppDriver
Private p_sBranchCd As String

Private p_orsMaster As Recordset
Private p_oClient As clsNeoClient
Private p_oReferd As clsNeoClient

Private p_sSQLMastr As String
Private p_sSQLBrows As String

Private p_bInitClas As Boolean
Private p_nEditMode As xeEditMode
Private p_aOthrInfo(80 To 86) As String
Private p_nTranStat As Integer
Private p_dApproved As Date

Public Event MasterRetrieved(ByVal Index As Variant)

'@@@@@@@@@@@@@@@@@@
'Properties here
Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get AppDriver() As clsAppDriver
   AppDriver = p_oAppDrivr
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_nEditMode
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal vValue As String)
   p_sBranchCd = vValue
End Property

Property Let TransStatus(ByVal Value As Integer)
   p_nTranStat = Value
End Property

Property Let Master(ByVal Index As Variant, ByVal fvValue As Variant)
   If p_bInitClas = False Then Exit Property
   If p_nEditMode = xeModeUnknown Then Exit Property
   
   Debug.Print Index & "»" & fvValue
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 0, "stransnox"
   Case 1, "dtransact"
      If IsDate(fvValue) Then
         p_orsMaster(Index) = CDate(fvValue)
      Else
         p_orsMaster(Index) = Null
      End If
      
      RaiseEvent MasterRetrieved(1)
   Case 2, "sclientid"
      p_orsMaster(Index) = fvValue
   Case 3, "smodelidx"
      p_orsMaster(Index) = fvValue
   Case 4, "scoloridx"
      p_orsMaster(Index) = fvValue
   Case 5, "sreferdby"
      p_orsMaster(Index) = fvValue
   Case 6, "dtargetxx"
      If IsDate(fvValue) Then
         p_orsMaster(Index) = CDate(fvValue)
      Else
         p_orsMaster(Index) = Null
      End If
      
      RaiseEvent MasterRetrieved(6)
      
   Case 7, "dfollowup"
      If IsDate(fvValue) Then
         p_orsMaster(Index) = CDate(fvValue)
      Else
         p_orsMaster(Index) = Null
      End If
      
      RaiseEvent MasterRetrieved(7)
   Case 8, "sremarksx"
      p_orsMaster(Index) = fvValue
   Case 9, "srelatnid"
      p_orsMaster(Index) = fvValue
   Case 10, "ssourcecd"
      p_orsMaster(Index) = fvValue
   Case 11, "ssourceno"
      p_orsMaster(Index) = fvValue
      
   Case 12, "ctranstat"
   Case 13, "smodified"
   Case 14, "dmodified"
   
   Case 80, "sclientnm"
      Call getCustomer(80, 2, fvValue, False)
   Case 81, "saddressx"
   Case 82, "smodelnme"
      Call getModel(82, 3, fvValue, False)
   Case 83, "scolornme"
      Call getColor(83, 4, fvValue, False)
   Case 84, "xreferdby"
      Call getCustomer(84, 5, fvValue, False)
   Case 85, "xaddressx"
   Case 86, "srelatnds"
      Call getRelation(86, 9, fvValue, False)
   End Select
End Property

Property Get Master(ByVal Index As Variant) As Variant
   If p_bInitClas = False Then Exit Property
   If p_nEditMode = xeModeUnknown Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 0, "stransnox"
      Master = p_orsMaster(Index)
   Case 1, "dtransact"
      Master = p_orsMaster(Index)
   Case 2, "sclientid"
      Master = p_orsMaster(Index)
   Case 3, "smodelidx"
      Master = p_orsMaster(Index)
   Case 4, "scoloridx"
      Master = p_orsMaster(Index)
   Case 5, "sreferdby"
      Master = p_orsMaster(Index)
   Case 6, "dtargetxx"
      Master = p_orsMaster(Index)
   Case 7, "dfollowup"
      Master = p_orsMaster(Index)
   Case 8, "sremarksx"
      Master = p_orsMaster(Index)
   Case 9, "srelatnid"
      Master = p_orsMaster(Index)
   Case 10, "ssourcecd"
      Master = p_orsMaster(Index)
   Case 11, "ssourceno"
      Master = p_orsMaster(Index)
   Case 12, "ctranstat"
   Case 13, "smodified"
   Case 14, "dmodified"
   
   Case 80, "sclientnm"
      If p_aOthrInfo(80) = "" Then
         If Trim(IFNull(p_orsMaster(2))) = "" Then
            p_aOthrInfo(80) = ""
         Else
            Call getCustomer(80, 2, p_orsMaster(2), True)
         End If
      End If
      Master = p_aOthrInfo(80)
   Case 81, "saddressx"
      If p_aOthrInfo(80) = "" Then
         If Trim(IFNull(p_orsMaster(2))) = "" Then
            p_aOthrInfo(81) = ""
         Else
            Call getCustomer(80, 2, p_orsMaster(2), True)
         End If
      End If
      Master = p_aOthrInfo(81)
   Case 82, "smodelnme"
      If p_aOthrInfo(82) = "" Then
         If Trim(IFNull(p_orsMaster(3))) = "" Then
            p_aOthrInfo(82) = ""
         Else
            Call getModel(82, 3, p_orsMaster(3), True)
         End If
      End If
      Master = p_aOthrInfo(82)
   Case 83, "scolornme"
      If p_aOthrInfo(83) = "" Then
         If Trim(IFNull(p_orsMaster(4))) = "" Then
            p_aOthrInfo(83) = ""
         Else
            Call getColor(83, 4, p_orsMaster(4), True)
         End If
      End If
      Master = p_aOthrInfo(83)

   Case 84, "xreferdby"
      If p_aOthrInfo(84) = "" Then
         If Trim(IFNull(p_orsMaster(5))) = "" Then
            p_aOthrInfo(84) = ""
         Else
            Call getReferredBy(84, 5, p_orsMaster(5), True)
         End If
      End If
      Master = p_aOthrInfo(84)
   Case 85, "xaddressx"
      If p_aOthrInfo(85) = "" Then
         If Trim(IFNull(p_orsMaster(5))) = "" Then
            p_aOthrInfo(85) = ""
         Else
            Call getReferredBy(84, 5, p_orsMaster(5), True)
         End If
      End If
      Master = p_aOthrInfo(85)
   
   Case 86, "srelatnds"
      If p_aOthrInfo(86) = "" Then
         If Trim(IFNull(p_orsMaster(9))) = "" Then
            p_aOthrInfo(86) = ""
         Else
            Call getRelation(86, 9, p_orsMaster(9), True)
         End If
      End If
      Master = p_aOthrInfo(86)
   End Select

endProc:
End Property

Function InitTransaction(Optional oAppDriver As Variant) As Boolean
   Dim lsSQL As String
   Dim lsCondition As String
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc

   'Perform the initialization if this is the first time this module was executed...
   If Not p_bInitClas Then
      If IsMissing(oAppDriver) Then
         If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
      Else
         If TypeName(oAppDriver) <> "clsAppDriver" Then GoTo endProc
         Set p_oAppDrivr = oAppDriver
      End If
   
      If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
      p_sSQLMastr = "SELECT " & _
                           "  a.sTransNox" & _
                           ", a.dTransact" & _
                           ", a.sClientID" & _
                           ", a.sModelIDx" & _
                           ", a.sColorIDx" & _
                           ", a.sReferdBy" & _
                           ", a.dTargetxx" & _
                           ", a.dFollowUp" & _
                           ", a.sRemarksx" & _
                           ", a.sRelatnID" & _
                           ", a.sSourceCd" & _
                           ", a.sSourceNo" & _
                           ", a.cTranStat" & _
                           ", a.sModified" & _
                           ", a.dModified" & _
                    " FROM MC_Referral a"
   
      p_sSQLBrows = "SELECT" & _
                           "  a.sTransNox" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNULL(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) sFullname" & _
                           ", a.dTransact" & _
                    " FROM MC_Referral a" & _
                        ", Client_Master b" & _
                    " WHERE a.sClientID = b.sClientID"
   
      If p_nTranStat <> xeStateUnknown Then
         If p_nTranStat > xeStateUnknown Then
            lsCondition = "("
            For lnCtr = 1 To Len(Trim(p_nTranStat))
               lsCondition = lsCondition & " cTranStat = " & _
                                 strParm(Mid(Trim(p_nTranStat), lnCtr, 1)) & " OR "
            Next
            lsCondition = Left(lsCondition, Len(Trim(lsCondition)) - 2) & ")"
         Else
            lsCondition = "a.cTranStat = " & strParm(p_nTranStat)
         End If
      End If
      
      If lsCondition <> "" Then
         p_sSQLBrows = AddCondition(p_sSQLBrows, lsCondition)
         p_sSQLMastr = AddCondition(p_sSQLMastr, lsCondition)
      End If
   
   End If
      
   Set p_oClient = New clsNeoClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
      
   Set p_oReferd = New clsNeoClient
   With p_oReferd
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
      
   p_bInitClas = True
   p_nEditMode = xeModeUnknown
   InitTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Function NewTransaction() As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "NewTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc


   'Perform the initialization if this is the first time this module was executed...
   If Not p_bInitClas Then Exit Function
   
   lsSQL = AddCondition(p_sSQLMastr, "0=1")
   Set p_orsMaster = New Recordset
   p_orsMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_orsMaster.ActiveConnection = Nothing
   
   With p_orsMaster
      .AddNew
      For lnCtr = 0 To .Fields.Count - 2
         Select Case Left(.Fields(lnCtr).Name, 1)
            Case "c"
               .Fields(lnCtr) = "0"
            Case "n"
               .Fields(lnCtr) = 0
            Case "d"
'               .Fields(lnCtr) = Null
            Case Else
               .Fields(lnCtr) = ""
         End Select
      Next
      .Fields("sTransNox") = GetNextCode("MC_Referral", "sTransNox", True, p_oAppDrivr.Connection, True, p_sBranchCd)
      .Fields("dTransact") = p_oAppDrivr.SysDate
   End With
   
   Call InitOthers
   
   p_nEditMode = xeModeAddNew
   NewTransaction = True

End Function

Function SearchMaster( _
         ByVal Index As Variant, _
         Optional fvValue As Variant = "" _
         ) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SearchMaster"
   'On Error GoTo errProc
   
   If p_bInitClas = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 80, "sclientnm"
      SearchMaster = getCustomer(80, 2, fvValue, False, True)
   Case 82, "smodelnme"
      SearchMaster = getModel(82, 3, fvValue, False, True)
   Case 83, "scolornme"
      SearchMaster = getColor(83, 4, fvValue, False, True)
   Case 84, "xreferdby"
      SearchMaster = getReferredBy(84, 5, fvValue, False, True)
   Case 86, "srelatnds"
      SearchMaster = getRelation(86, 9, fvValue, False, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & Index _
                        & ", " & fvValue & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "OpenTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   'Perform the initialization if this is the first time this module was executed...
   If Not p_bInitClas Then Exit Function
   
   lsSQL = AddCondition(p_sSQLMastr, "a.sTransNox = " & strParm(sTransNo))
   Set p_orsMaster = New Recordset
   p_orsMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set p_orsMaster.ActiveConnection = Nothing
   
   If p_orsMaster.EOF Then
      p_nEditMode = xeModeUnknown
      Exit Function
   End If
      
   Call InitOthers
   
   p_nEditMode = xeModeUpdate
   OpenTransaction = True

End Function

Function SearchTransaction(ByVal Value As String, Optional ByVal isCode As Boolean = False) As Boolean
   Dim lsSQL As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim loRS As Recordset
   Dim lsOldProc As String
   
   lsOldProc = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUpdate Or p_nEditMode = xeModeReady Then
      If isCode Then
         If Value = p_orsMaster("sTransNox") Then
            MsgBox "Record seems to be loaded! Please check your entry..."
            GoTo endProc
         End If
      End If
   End If
   
   If isCode Then
      lsSQL = AddCondition(p_sSQLBrows, "a.sTransNox LIKE " & strParm("%" & Value))
   Else
      If Value = "" Then
         lsSQL = p_sSQLBrows
      Else
         lasSelected = GetSplitedName(Value)
         lsCondition = "b.sLastName LIKE " & strParm(lasSelected(0) & "%") & _
                  " AND b.sFrstName LIKE " & strParm(IIf(UBound(lasSelected) > 0, lasSelected(1) & "%", "")) & _
                  " AND b.sMiddName LIKE " & strParm(IIf(UBound(lasSelected) > 1, lasSelected(2) & "%", ""))
         lsSQL = AddCondition(p_sSQLBrows, lsCondition)
      End If
   End If
            
   'If the application is not the Telemarketing System then filter search by branch
   'kalyptus - 2016.03.05 02:01pm
   'TODO: Determine the name of the System to be use here...
   If UCase(p_oAppDrivr.ProductID) <> "LRTRACKR" Then
      'Don't allow user to view loans if not manager or engineer
      If (p_oAppDrivr.UserLevel And (xeManager + xeEngineer + xeSysAdmin)) = 0 Then
          GoTo endProc
      End If
   
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(p_sBranchCd & "%"))
   End If
   
   'Search and load encoded entries
   Set loRS = New Recordset
   Debug.Print lsSQL
   With loRS
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         GoTo endProc
      ElseIf .RecordCount = 1 Then
         SearchTransaction = OpenTransaction(.Fields("sTransNox"))
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, loRS)
         If lsSQL <> "" Then
            lasSelected = Split(lsSQL, "»")
            SearchTransaction = OpenTransaction(lasSelected(0))
         End If
      End If
   End With
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "SaveTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
   
   If Not isEntryOK Then GoTo endProc
         
   'Save Client Info
   If Not p_oClient.SaveClient Then
      MsgBox "Client Info was not save!" & vbCrLf & _
                   "Please make sure to re-enter client info.", vbCritical, "Warning"
      Exit Function
   End If
      
   If Not p_oReferd.SaveClient Then
      MsgBox "Client Info was not save!" & vbCrLf & _
                   "Please make sure to re-enter client info.", vbCritical, "Warning"
      Exit Function
   End If
      
      
   p_orsMaster("sClientID") = p_oClient.Master("sClientID")
   p_orsMaster("sReferdBy") = p_oReferd.Master("sClientID")
      
   If p_nEditMode = xeModeAddNew Then
      p_orsMaster("sTransnox") = GetNextCode("MC_Referral", _
                  "sTransNox", _
                  True, _
                  p_oAppDrivr.Connection, _
                  True, _
                  p_sBranchCd)
      
      lsSQL = ADO2SQL(p_orsMaster, "MC_Referral", , p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
   Else
      lsSQL = ADO2SQL(p_orsMaster, "MC_Referral", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
   End If
      
   If lsSQL <> "" Then
      p_oAppDrivr.BeginTrans
      Debug.Print lsSQL
      p_oAppDrivr.Execute lsSQL, "MC_Referral", Left(p_orsMaster("sTransNox"), 4)
      
      p_oAppDrivr.CommitTrans
   End If
   
   SaveTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

'Used for indicating inquiry is scheduled for call by telemarketing agent
Function CloseTransaction() As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim loPayMiscx As clsPayMisc
   Dim ldValApprv As Date
   Dim p_dApproved As Date
   
      
   lsOldProc = "CloseTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
      
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or _
      p_orsMaster("cTranStat") = xeStateCancelled Or _
      p_orsMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/disapproved. You are no longer allowed to disapproved the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   p_oAppDrivr.BeginTrans
         
   p_orsMaster("cTranStat") = xeStateClosed
   p_orsMaster("sApproved") = p_oAppDrivr.UserID
   p_orsMaster("dApproved") = p_dApproved
   
   lsSQL = ADO2SQL(p_orsMaster, "MC_Referral", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
   
   p_oAppDrivr.Execute lsSQL, "MC_Referral", , Left(p_orsMaster("sTransNox"), 4)
   
   p_oAppDrivr.CommitTrans
   
   CloseTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

'Used for indicating purchase
Function PostTransaction() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim p_dApproved As Date
   
   lsOldProc = "PostTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
      
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or _
      p_orsMaster("cTranStat") = xeStateCancelled Or _
      p_orsMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/disapproved. You are no longer allowed to disapproved the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   p_oAppDrivr.BeginTrans
         
   p_orsMaster("cTranStat") = xeStatePosted
   p_orsMaster("sApproved") = p_oAppDrivr.UserID
   p_orsMaster("dApproved") = p_dApproved

   lsSQL = ADO2SQL(p_orsMaster, "MC_Referral", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)

   p_oAppDrivr.Execute lsSQL, "MC_Referral", , Left(p_orsMaster("sTransNox"), 4)
   
   p_oAppDrivr.CommitTrans
      
   PostTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

'Used for indicating that customer declined
Function CancelTransaction() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "CancelTransaction"
   'On Error GoTo errProc
      
   If p_nEditMode = xeModeUnknown Then GoTo endProc
      
   'Test if record is posted
   If p_orsMaster("cTranStat") = xeStatePosted Or _
      p_orsMaster("cTranStat") = xeStateCancelled Or _
      p_orsMaster("cTranStat") = xeStateUnknown Then
      MsgBox "Record was posted/disapproved. You are no longer allowed to disapproved the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_oAppDrivr.BeginTrans

   p_orsMaster("cTranStat") = xeStateCancelled
   p_orsMaster("sApproved") = p_oAppDrivr.UserID
   p_orsMaster("dApproved") = p_dApproved

   lsSQL = ADO2SQL(p_orsMaster, "MC_Referral", "sTransNox = " & strParm(p_orsMaster("sTransNox")), p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)

   p_oAppDrivr.Execute lsSQL, "MC_Referral", , Left(p_orsMaster("sTransNox"), 4)
      
   p_oAppDrivr.CommitTrans
   CancelTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function isEntryOK() As Boolean
   If p_orsMaster("cTranStat") = "2" Then
      MsgBox "Inquiry was posted! Posted application are no longer allowed to update!", vbCritical, "MC Referral Entry Validation"
      Exit Function
   End If
   
   If p_orsMaster("cTranStat") = "3" Then
      MsgBox "Inquiry was cancelled! Cancelled application are no longer allowed to update!", vbCritical, "MC Referral Entry Validation"
      Exit Function
   End If
   
   If IFNull(p_orsMaster("sClientID")) = "" Then
      MsgBox "Invalid Client detected!  Please check client...", vbCritical, "MC Referral Entry Validation"
      Exit Function
   End If
   
   If IFNull(p_orsMaster("sReferdBy")) = "" Then
      MsgBox "Invalid Referral ID detected!  Please check referral info...", vbCritical, "MC Referral Entry Validation"
      Exit Function
   End If
   
   If IFNull(p_orsMaster("sModelIDx")) = "" Then
      MsgBox "Invalid Model detected!  Please Model employee...", vbCritical, "MC Referral Entry Validation"
      Exit Function
   End If
   
   isEntryOK = True
End Function

Private Sub InitOthers()
   Dim lnCtr As Integer
   
   For lnCtr = LBound(p_aOthrInfo) To UBound(p_aOthrInfo)
      p_aOthrInfo(lnCtr) = ""
   Next

   Call p_oClient.InitClient
   Call p_oReferd.InitClient

End Sub

Private Function getCustomer( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
      'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   If fbCode Then
      If loClient.OpenClient(fsValue) Then
         Set p_oClient = loClient
      Else
         GoTo endWithClear
      End If
   Else
      
      lbExist = loClient.SearchClient(fsValue, Not fbSearch)
      
      If Not lbExist And fsValue <> "" Then
         lasName = GetSplitedName(fsValue)
         loClient.Master("sLastName") = lasName(0)
         loClient.Master("sFrstName") = lasName(1)
      End If

      If loClient.getClient Then
         Set p_oClient = loClient
      End If
   End If
   
   p_orsMaster("sClientID") = p_oClient.Master("sClientID")
   If p_oClient.Master("cMCClient") <> "1" Then p_oClient.Master("cMCClient") = "3"
   
   If p_orsMaster("cPurcType") = "1" Then
      If p_oClient.Master("cLRClient") <> "1" Then p_oClient.Master("cLRClient") = "3"
   End If
   
   p_aOthrInfo(80) = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
   p_aOthrInfo(81) = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")
   
   getCustomer = True
   
endProc:
   If Not fbCode Then
      RaiseEvent MasterRetrieved(80)
      RaiseEvent MasterRetrieved(81)
   End If
   
   Exit Function
endWithClear:
   p_orsMaster("sClientID") = ""
   Set p_oClient = loClient
   p_aOthrInfo(80) = ""
   p_aOthrInfo(81) = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & fsValue _
                        & ", " & fbSearch & " )"
End Function

Private Function getModel( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lasMaster() As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim lsOldProc As String
         
   lsOldProc = pxeMODULENAME & "." & "getModel(" & fnRow & ", " & fsValue & ")"
   'On Error GoTo errProc
      
   'Make sure that fetched fields are new entries
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
            "  sModelIDx" & _
            ", sModelNme" & _
         " FROM MC_Model" & _
         " WHERE cRecdStat <> '0'"
                 
   If fbCode = False Then
      If fbSearch Then
         lsCondition = "sModelNme LIKE " & strParm(fsValue & "%")
      Else
         lsCondition = "sModelNme = " & strParm(fsValue)
      End If
      lsSQL = AddCondition(lsSQL, lsCondition)
   Else
      lsSQL = AddCondition(lsSQL, "sModelIDx = " & strParm(Trim(fsValue)))
   End If
   
   'Search and load encoded entries
   Set loRS = New Recordset
   Debug.Print lsSQL
   With loRS
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         p_orsMaster(fnID) = ""
         p_aOthrInfo(fnRow) = ""
      ElseIf .RecordCount = 1 Then
         p_orsMaster(fnID) = loRS("sModelIDx")
         p_aOthrInfo(fnRow) = loRS("sModelNme")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, loRS, "sModelIDx»sModelNme", "ID»Model")
         If lsSQL <> "" Then
            lasMaster = Split(lsSQL, "»")
            p_orsMaster(fnID) = lasMaster(0)
            p_aOthrInfo(fnRow) = lasMaster(1)
         End If
      End If
      RaiseEvent MasterRetrieved(fnRow)
   End With
      
   getModel = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function getColor( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lasMaster() As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim lsOldProc As String
         
   lsOldProc = pxeMODULENAME & "." & "getColor(" & fnRow & ", " & fsValue & ")"
   'On Error GoTo errProc
      
   'Make sure that fetched fields are new entries
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
            "  sColorIDx" & _
            ", sColorNme" & _
         " FROM Color" & _
         " WHERE cRecdStat <> '0'"
                 
   If fbCode = False Then
      If fbSearch Then
         lsCondition = "sColorNme LIKE " & strParm(fsValue & "%")
      Else
         lsCondition = "sColorNme = " & strParm(fsValue)
      End If
      lsSQL = AddCondition(lsSQL, lsCondition)
   Else
      lsSQL = AddCondition(lsSQL, "sColorIDx = " & strParm(Trim(fsValue)))
   End If
   
   'Search and load encoded entries
   Set loRS = New Recordset
   Debug.Print lsSQL
   With loRS
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         p_orsMaster(fnID) = ""
         p_aOthrInfo(fnRow) = ""
      ElseIf .RecordCount = 1 Then
         p_orsMaster(fnID) = loRS("sColorIDx")
         p_aOthrInfo(fnRow) = loRS("sColorNme")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, loRS, "sColorIDx»sColorNme", "ID»Color")
         If lsSQL <> "" Then
            lasMaster = Split(lsSQL, "»")
            p_orsMaster(fnID) = lasMaster(0)
            p_aOthrInfo(fnRow) = lasMaster(1)
         End If
      End If
      RaiseEvent MasterRetrieved(fnRow)
   End With
      
   getColor = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function getReferredBy( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getReferredBy"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
      'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   If fbCode Then
      If loClient.OpenClient(fsValue) Then
         Set p_oReferd = loClient
      Else
         GoTo endWithClear
      End If
   Else
      
      lbExist = loClient.SearchClient(fsValue, Not fbSearch)
      
      If Not lbExist And fsValue <> "" Then
         lasName = GetSplitedName(fsValue)
         loClient.Master("sLastName") = lasName(0)
         loClient.Master("sFrstName") = lasName(1)
      End If

      If loClient.getClient Then
         Set p_oReferd = loClient
      End If
   End If
   
   p_orsMaster("sReferdBy") = p_oReferd.Master("sClientID")
   If p_oReferd.Master("cMCClient") <> "1" Then p_oReferd.Master("cMCClient") = "3"
   
   If p_orsMaster("cPurcType") = "1" Then
      If p_oReferd.Master("cLRClient") <> "1" Then p_oReferd.Master("cLRClient") = "3"
   End If
   
   p_aOthrInfo(84) = p_oReferd.Master("sLastName") + ", " + p_oReferd.Master("sFrstName") + " " + Trim(p_oReferd.Master("sSuffixNm")) + IIf(Trim(p_oReferd.Master("sSuffixNm")) = "", "", " ") + p_oReferd.Master("sMiddName")
   p_aOthrInfo(85) = IIf(Trim(p_oReferd.Master("sHouseNox")) = "", "", p_oReferd.Master("sHouseNox") & " ") & p_oReferd.Master("sAddressx") & ", " & p_oReferd.Master("sTownName")
   
   getReferredBy = True
   
endProc:
   If Not fbCode Then
      RaiseEvent MasterRetrieved(84)
      RaiseEvent MasterRetrieved(85)
   End If
   
   Exit Function
endWithClear:
   p_orsMaster("sReferdBy") = ""
   Set p_oReferd = loClient
   p_aOthrInfo(84) = ""
   p_aOthrInfo(85) = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & fsValue _
                        & ", " & fbSearch & " )"
End Function

Private Function getRelation( _
   ByVal fnRow As Integer, _
   ByVal fnID As Integer, _
   ByVal fsValue As String, _
   Optional ByVal fbCode As Boolean = False, _
   Optional ByVal fbSearch As Boolean = False) As Boolean
   
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lasMaster() As String
   Dim lasSelected() As String
   Dim lsCondition As String
   Dim lsOldProc As String
         
   lsOldProc = pxeMODULENAME & "." & "getModel(" & fnRow & ", " & fsValue & ")"
   'On Error GoTo errProc
      
   'Make sure that fetched fields are new entries
   If Not (fbSearch Or fbCode) Then
      If Trim(fsValue) = Trim(p_aOthrInfo(fnRow)) Then GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
            "  sRelatnID" & _
            ", sRelatnDs" & _
         " FROM Relation" & _
         " WHERE cRecdStat <> '0'"
                 
   If fbCode = False Then
      If fbSearch Then
         lsCondition = "sRelatnDs LIKE " & strParm(fsValue & "%")
      Else
         lsCondition = "sRelatnDs = " & strParm(fsValue)
      End If
      lsSQL = AddCondition(lsSQL, lsCondition)
   Else
      lsSQL = AddCondition(lsSQL, "sRelatnID = " & strParm(Trim(fsValue)))
   End If
   
   'Search and load encoded entries
   Set loRS = New Recordset
   Debug.Print lsSQL
   With loRS
      .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If .EOF Then
         p_orsMaster(fnID) = ""
         p_aOthrInfo(fnRow) = ""
      ElseIf .RecordCount = 1 Then
         p_orsMaster(fnID) = loRS("sRelatnID")
         p_aOthrInfo(fnRow) = loRS("sRelatnDs")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, loRS, "sRelatnID»sRelatnDs", "ID»Relation")
         If lsSQL <> "" Then
            lasMaster = Split(lsSQL, "»")
            p_orsMaster(fnID) = lasMaster(0)
            p_aOthrInfo(fnRow) = lasMaster(1)
         End If
      End If
      RaiseEvent MasterRetrieved(fnRow)
   End With
      
   getRelation = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function


Private Sub ShowError(ByVal lsOldProc As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsOldProc, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

Private Sub Class_Initialize()
   p_nTranStat = xeStateUnknown
End Sub



