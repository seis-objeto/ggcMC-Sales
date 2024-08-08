VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Releasing of Commission object
'
' Copyright 2015 and beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 05/15/2015 11:49 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsCommission"
Private Const pxeTableNme As String = "MC_SO_Agent"

Private p_oAppDrivr As clsAppDriver

Private p_sBranchCd As String
Private p_sParentxx As String

Private p_oMaster As Recordset
Private p_sMasterSQL As String
Private p_sBrowseSQL As String
Private p_cTranStat As xeTransactionStatus
Private p_nEditMode As xeEditMode

Private pbInitTran As Boolean
Private pbShowMsg As Boolean
Private pnCtr As Integer

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let Status(ByVal Value As xeTransactionStatus)
   Select Case Value
   Case xeStateOpen, xeStateClosed, xeStateCancelled, xeStateUnknown
      p_cTranStat = Value
   End Select
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Master = p_oMaster(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
End Property

Property Let Parent(ByVal Value As String)
   p_sParentxx = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQLBrows As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   p_sMasterSQL = "SELECT a.sReleasNo" & _
                     ", a.sTransNox" & _
                     ", b.sCompnyNm xAgentNme" & _
                     ", CONCAT(b.sAddressx, ', ', c.sTownName) xAddressA" & _
                     ", h.sCompnyNm xCustomer" & _
                     ", CONCAT(h.sAddressx, ', ', i.sTownName) xAddressC" & _
                     ", d.sDRNoxxxx" & _
                     ", g.sModelNme" & _
                     ", f.sEngineNo" & _
                     ", a.nCommAmtx" & _
                     ", a.sReferNox" & _
                     ", a.cReleased" & _
                     ", a.sReleased" & _
                     ", a.dReleased" & _
                     ", a.dModified" & _
                     ", a.sAgentIDx" & _
                     ", a.cGCRefrlx"
                     
   p_sMasterSQL = p_sMasterSQL & _
                  " FROM " & pxeTableNme & " a" & _
                     ", Client_Master b" & _
                     ", TownCity c" & _
                     ", MC_SO_Master d" & _
                     ", MC_SO_Detail e" & _
                     ", MC_Serial f" & _
                     ", MC_Model g" & _
                     ", Client_Master h" & _
                     ", TownCity i" & _
                  " WHERE a.sAgentIDx = b.sClientID" & _
                     " AND b.sTownIDxx = c.sTownIDxx" & _
                     " AND a.sTransNox = d.sTransNox" & _
                     " AND d.sTransNox = e.sTransNox" & _
                     " AND e.sSerialID = f.sSerialID" & _
                     " AND f.sModelIDx = g.sModelIDx" & _
                     " AND d.sClientID = h.sClientID" & _
                     " AND h.sTownIDxx = i.sTownIDxx"
                  
   
   Set p_oMaster = New Recordset
   With p_oMaster
      .Open AddCondition(p_sMasterSQL, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      .AddNew
      Call InitCommission
   End With
   
   p_sBrowseSQL = "SELECT a.sTransNox" & _
                     ", b.sCompnyNm xAgentNme" & _
                     ", d.sCompnyNm xCustomer" & _
                     ", c.sDRNoxxxx" & _
                  " FROM " & pxeTableNme & " a" & _
                     ", Client_Master b" & _
                     ", MC_SO_Master c" & _
                     ", Client_Master d" & _
                  " WHERE a.sAgentIDx = b.sClientID" & _
                     " AND a.sTransNox = c.sTransNox" & _
                     " AND c.sClientID = d.sClientID"
   
   p_nEditMode = xeModeReady
   pbInitTran = True
   InitTransaction = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Sub InitCommission()
   With p_oMaster
      .Fields("sReleasNo") = ""
      .Fields("sTransNox") = ""
      .Fields("xAgentNme") = ""
      .Fields("xCustomer") = ""
      .Fields("sDRNoxxxx") = ""
      .Fields("sModelNme") = ""
      .Fields("sEngineNo") = ""
      .Fields("nCommAmtx") = 0
      .Fields("sReferNox") = ""
      .Fields("cGCRefrlx") = xeNo
      .Fields("cReleased") = xeNo
      .Fields("sReleased") = ""
      .Fields("dReleased") = Date
      .Fields("dModified") = Date
   End With
End Sub

'Function NewTransaction() As Boolean
'   Dim lsProcName As String
'
'   lsProcName = "NewTransaction"
'   'On Error GoTo errProc
'
'   If Not pbInitTran Then GoTo endProc
'
'   If p_nEditMode <> xeModeReady Then
'      Call showMessage("Transaction is in update mode!" & vbCrLf & vbCrLf & _
'            "All modification will be lost!", , "Warning")
'   End If
'
'   p_oMaster.AddNew
'   Call InitCommission
'   p_nEditMode = xeModeAddNew
'
'   NewTransaction = True
'
'endProc:
'   Exit Function
'errProc:
'    ShowError lsProcName & "( " & " )"
'End Function
'

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnAppvRights As Integer, lnAllowRights As Integer
   Dim lsAppvID As String, lsAppvName As String

   lsProcName = "CancelTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If p_nEditMode <> xeModeReady Then
      showMessage "No transaction is loaded for releasing!", , "Warning"
      GoTo endProc
   End If
   
   If pbShowMsg Then
      If showMessage("Are you sure you want to Cancel the Commission?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
      
      p_oAppDrivr.BeginTrans
   End If
   
   lnAllowRights = 232 'p_oAppDrivr.getConfiguration("MCGAway")
   If lnAllowRights And p_oAppDrivr.UserLevel = 0 Then
      If Not pbShowMsg Then GoTo endProc
      
      If showMessage("User Doesn't Have the Right to Add New Giveaway!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes Then
         If GetApproval(p_oAppDrivr, _
               lnAppvRights, _
               lsAppvID, _
               lsAppvName, _
               p_oAppDrivr.MenuName) = False Then GoTo endProc
         
         If lnAllowRights And (lnAppvRights = 0) Then
            showMessage "Approving Officer Has no Right to Add New Giveaway!!!" & vbCrLf & _
               "Request can not be granted!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End If
   
   lsSQL = "UPDATE " & pxeTableNme & _
            " SET cReleased = " & strParm(xeStateCancelled) & _
               ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
            " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
            
   If p_oAppDrivr.Execute(lsSQL, pxeTableNme, p_sBranchCd) = 0 Then
      MsgBox "Unable to Release Agent Commission!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   ' check if commission was already released
   If p_oMaster("cReleased") = xeYes Then
      lsSQL = "UPDATE MC_Agent" & _
               " SET nNoOfCust = nNoOfCust - 1" & _
                  ", nCommissn = nCommissn - " & p_oMaster("nCommAmtx") & _
               " WHERE sClientID = " & strParm(p_oMaster("sAgentIDx"))
               
      If p_oAppDrivr.Execute(lsSQL, "MC_Agent") = 0 Then
         MsgBox "Unable to Update Agent Information!!!", vbCritical, "Warning"
         GoTo endWithRoll
      End If
   End If
   
   If pbShowMsg Then
      p_oAppDrivr.CommitTrans
   End If
   
   CancelTransaction = True

endProc:
   Exit Function
endWithRoll:
   If pbShowMsg Then
      p_oAppDrivr.RollbackTrans
   End If
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "OpenTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If p_nEditMode <> xeModeReady Then
      Call showMessage("Transaction is in update mode!" & vbCrLf & vbCrLf & _
            "All modification will be lost!", , "Warning")
   End If
   
   Set p_oMaster = New Recordset
   With p_oMaster
      lsSQL = AddCondition(p_sMasterSQL, "a.sTransNox = " & strParm(sTransNo))
      .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      
      If .EOF Then
         .AddNew
         Call InitCommission
         p_nEditMode = xeModeUnknown
         GoTo endProc
      End If
      
      ' Assign default transaction number
      .Fields("sReleasNo") = GetNextCode(pxeTableNme, "sReleasNo", True, p_oAppDrivr.Connection, True, p_sBranchCd)
   End With
      
   OpenTransaction = True
   p_nEditMode = xeModeReady

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant = "", _
                           Optional bByCode As Variant = False) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If sSearch <> "" Then
      If bByCode Then
         If p_oMaster("sDRNoxxxx") = sSearch Then
            SearchTransaction = True
            GoTo endProc
         End If
      Else
         If p_oMaster("xAgentNme") = sSearch Then
            SearchTransaction = True
            GoTo endProc
         End If
      End If
   End If
      
   If bByCode Then
      lsSQL = AddCondition(p_sBrowseSQL, "c.sDRNoxxxx LIKE " & strParm(sSearch & "%"))
   Else
      lsSQL = AddCondition(p_sBrowseSQL, "b.sCompnyNm LIKE " & strParm(sSearch & "%"))
   End If
   
   If p_cTranStat <> xeStateUnknown Then
      lsSQL = AddCondition(lsSQL, "a.cReleased = " & strParm(p_cTranStat))
   End If
   Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   With loRS
      If .EOF Then GoTo endProc
      
      If .RecordCount > 1 Then
         lsSQL = KwikBrowse(p_oAppDrivr, loRS, , "Trans #»Agent Name»Customer»DR #")
         
         If lsSQL = "" Then GoTo endProc
         
         lasMaster = Split(lsSQL, "»")
         
         Call .Find("sTransNox = " & strParm(lasMaster(0)))
      End If
      SearchTransaction = OpenTransaction(.Fields("sTransNox"))
   End With

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

'Function UpdateTransaction() As Boolean
'   If pbInitTran = False Then Exit Function
'
'   If p_nEditMode = xeModeUnknown Then
'      Call showMessage("No transaction is loaded for modification!", , "Warning")
'      GoTo endProc
'   ElseIf p_nEditMode = xeModeUpdate Then
'      Call showMessage("Transaction is in not save yet!" & vbCrLf & vbCrLf & _
'            "Update is not allowed!", , "Warning")
'      GoTo endProc
'   End If
'
'   p_nEditMode = xeModeUpdate
'   UpdateTransaction = True
'End Function

Function ReleasedCommission(ByVal sReferNox As String, ByVal dReleased As Date) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "ReleasedCommission"
   'On Error GoTo errProc
   
   If p_nEditMode <> xeModeReady Then
      showMessage "No transaction is loaded for releasing!", , "Warning"
      GoTo endProc
   End If
   
   'Mac 2018-05-09
   'check if it he is a GCardHolder
   If p_oMaster("cGCRefrlx") = xeYes Then
      showMessage "Unable to realease commision of a G-Card Holder." & vbCrLf & _
                  "You may use the G-Card System to release its commision by Points.", , "Warning"
      GoTo endProc
   End If
   
   If pbShowMsg Then
      If showMessage("Are you sure you want to Release the Commission?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
      p_oAppDrivr.BeginTrans
   End If
   
   lsSQL = "UPDATE " & pxeTableNme & _
            " SET sReleasNo = " & strParm(GetNextCode(pxeTableNme, "sReleasNo", True, _
                  p_oAppDrivr.Connection, True, p_sBranchCd)) & _
               ", sReferNox = " & strParm(sReferNox) & _
               ", dReleased = " & dateParm(dReleased) & _
               ", cReleased = " & strParm(xeYes) & _
               ", sReleased = " & strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
            " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
            
   If p_oAppDrivr.Execute(lsSQL, pxeTableNme, p_sBranchCd) = 0 Then
      MsgBox "Unable to Release Agent Commission!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
      
   lsSQL = "UPDATE MC_Agent" & _
            " SET nNoOfCust = nNoOfCust + 1" & _
               ", nCommissn = nCommissn + " & p_oMaster("nCommAmtx") & _
            " WHERE sClientID = " & strParm(p_oMaster("sAgentIDx"))
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Agent") = 0 Then
      MsgBox "Unable to Update Agent Information!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   If pbShowMsg Then
      p_oAppDrivr.CommitTrans
   End If
   
   ReleasedCommission = True
   
endProc:
   Exit Function
endWithRoll:
   If pbShowMsg Then
      p_oAppDrivr.RollbackTrans
      GoTo endProc
   End If
errProc:
    ShowError lsProcName & "( " & sReferNox & _
                           ", " & dReleased & " )"
End Function

Function UnReleasedCommission() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "UnReleasedCommission"
   'On Error GoTo errProc
   
   If p_nEditMode <> xeModeReady Then
      showMessage "No transaction is loaded for unreleasing!", , "Warning"
      GoTo endProc
   End If
   
   If pbShowMsg Then
      If showMessage("Are you sure you want to Unrelease the Commission?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
      
      p_oAppDrivr.BeginTrans
   End If
   
   If p_oMaster("cReleased") <> xeYes Then
      Call showMessage("Commission is not yet released!" & vbCrLf & _
               "Please verify your entry then try again!", vbInformation, "Warning")
      GoTo endProc
   End If
   
   If DateDiff("h", p_oMaster("dModified"), p_oAppDrivr.ServerDate) > 24 Then
      ' Unreleased of commission is only allowed with in 24 hours
      showMessage "Unrealesed of commission is only allowed" & vbCrLf & _
         "Within 24 hours after the tagging of the released!", , "Warning"
      GoTo endProc
   End If
   
   lsSQL = "UPDATE " & pxeTableNme & _
            " SET sReferNox = " & strParm("") & _
               ", cReleased = " & strParm(xeNo) & _
               ", sReleased = " & strParm("") & _
               ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
            " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
            
   If p_oAppDrivr.Execute(lsSQL, pxeTableNme, p_sBranchCd) = 0 Then
      MsgBox "Unable to Unrelease Agent Commission!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
      
   lsSQL = "UPDATE MC_Agent" & _
            " SET nNoOfCust = nNoOfCust - 1" & _
               ", nCommissn = nCommissn - " & p_oMaster("nCommAmtx") & _
            " WHERE sClientID = " & strParm(p_oMaster("sAgentIDx"))
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Agent") = 0 Then
      MsgBox "Unable to Update Agent Information!!!", vbCritical, "Warning"
      GoTo endWithRoll
   End If
   
   If pbShowMsg Then
      p_oAppDrivr.CommitTrans
   End If
   
   UnReleasedCommission = True
   
endProc:
   Exit Function
endWithRoll:
   If pbShowMsg Then
      p_oAppDrivr.RollbackTrans
   End If
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function showMessage(lsPrompt As String, _
      Optional lvStyle As Variant = vbOKOnly, _
      Optional lsTitle As String = "") As VbMsgBoxResult
   If pbShowMsg Then
      showMessage = MsgBox(lsPrompt, lvStyle, lsTitle)
   End If
End Function

Private Sub Class_Initialize()
   pbShowMsg = True
End Sub
