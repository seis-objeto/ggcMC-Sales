VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCWarranty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
'     Motorcycle Warranty Pullout Object
'
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
'  XerSys [ 04/24/2006 12:06 pm ]
'     Start creating this object.
'  XerSys [ 04/25/2006 08:50 am ]
'     Continue creating this object...
'  Kalyptus [ 12/11/2006 03:28 pm]
'     Implements the BugTracker
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oRecMaster As Recordset

Private p_sBranchCd As String
Private p_cTranStat As xeTransactionStatus
Private p_sFullName As String
Private p_sAddressx As String
Private p_sSerialID As String
Private p_sActQuery As String
Private p_nEditMode As xeEditMode

Private pbInitTran As Boolean
Private pbRecExist As Boolean
Private pnCtr As Integer

Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "xfullname"
      Master = p_sFullName
   Case 3, "xaddressx"
      Master = p_sAddressx
   Case Else
      Master = p_oRecMaster(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 2, "xfullname"
      getClient Value
   Case 3, "xaddressx"
      p_sAddressx = Value
   Case Else
      p_oRecMaster(Index) = Value
   End Select
End Property

Property Let TransStatus(ByVal Value As xeTransactionStatus)
   p_cTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String
    
   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   lsOldProc = "InitTransaction"
   'On Error Goto errProc
   InitTransaction = False
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   p_sActQuery = "SELECT" & _
                     "  a.sTransNox" & _
                     ", a.dTransact" & _
                     ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                     ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) AS xAddressx" & _
                     ", e.sEngineNo" & _
                     ", e.sFrameNox" & _
                     ", f.sModelNme" & _
                     ", g.dPurchase" & _
                     ", a.sRemarksx" & _
                     ", a.sApproved" & _
                     ", a.dReleased" & _
                     ", a.cTranStat" & _
                     ", a.sClientID" & _
                     ", a.sSerialID" & _
                     ", e.sMCInvIDx"
   p_sActQuery = p_sActQuery & _
                  " FROM MC_Warranty a" & _
                     ", Client_Master b" & _
                     ", TownCity c" & _
                     ", Province d" & _
                     ", MC_Serial e" & _
                     ", MC_Model f" & _
                     ", MC_Registration g" & _
                  " WHERE a.sClientID = b.sClientID" & _
                     " AND b.sTownIDxx = c.sTownIDxx" & _
                     " AND c.sProvIDxx = d.sProvIDxx" & _
                     " AND a.sSerialID = e.sSerialID" & _
                     " AND e.sModelIDx = f.sModelIDx" & _
                     " AND a.sSerialID = g.sSerialID" & _
                     " AND a.sClientID = g.sClientID" & _
                  " ORDER BY g.dPurchase"
                  
   Set p_oRecMaster = New Recordset
   p_oRecMaster.Open AddCondition(p_sActQuery, "0 = 1"), p_oAppDrivr.Connection, _
                        adOpenStatic, adLockOptimistic, adCmdText
   
   With p_oRecMaster
      Set .ActiveConnection = Nothing
      .AddNew
      .Fields("sTransNox") = GetNextCode("MC_Warranty", "sTransNox", True, _
                     p_oAppDrivr.Connection, True, p_sBranchCd)
      .Fields("dTransact") = Format(p_oAppDrivr.ServerDate, "mm/dd/yyyy")
      .Fields("cTranStat") = xeStateOpen
      .Fields("sEngineNo") = ""
      .Fields("sFrameNox") = ""
      .Fields("sModelNme") = ""
      .Fields("dPurchase") = .Fields("dTransact")
      .Fields("sRemarksx") = ""
      .Fields("sApproved") = ""
      .Fields("sClientID") = ""
      .Fields("sSerialID") = ""
      .Fields("sMCInvIDx") = ""
      
      p_sFullName = ""
      p_sAddressx = ""
      p_sSerialID = ""
      p_cTranStat = xeStateOpen
   End With
   p_nEditMode = xeModeAddNew
   
   pbRecExist = False
   pbInitTran = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "NewTransaction"
   'On Error Goto errProc

   If pbInitTran = False Then Exit Function
   NewTransaction = InitTransaction

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim loMCTrans As clsMotorcycle
   Dim lsOldProc As String, lsSQL As String
   Dim lsApproved As String
   Dim lnRow As Long

   lsOldProc = "SaveTransaction"
   'On Error Goto errProc
   SaveTransaction = False

   If pbInitTran = False Or _
         (pbRecExist = False And p_nEditMode = xeModeReady) Then GoTo endProc

   lsApproved = p_oAppDrivr.UserID
   If HasRights2Save(lsApproved) = False Then GoTo endProc
   
   With p_oRecMaster
      If .Fields("sClientID") = Empty Then
         MsgBox "Invalid Client Info Detected!!!" & vbCrLf & _
               "Verify your Entries then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If

      If .Fields("sSerialID") = Empty Then
         MsgBox "Invalid Motorcycle Detected!!!" & vbCrLf & _
               "Verify your Entries then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If

      p_oAppDrivr.BeginTrans
      If p_nEditMode = xeModeAddNew Then
         lsSQL = "INSERT INTO MC_Warranty " & _
                     "( sTransNox" & _
                     ", dTransact" & _
                     ", sClientID" & _
                     ", sSerialID" & _
                     ", sRemarksx" & _
                     ", sApproved" & _
                     ", cTranStat" & _
                     ", sModified" & _
                     ", dModified" & _
                  " ) VALUES ( " & _
                     strParm(.Fields("sTransNox")) & _
                     ", " & dateParm(.Fields("dTransact")) & _
                     ", " & strParm(.Fields("sClientID")) & _
                     ", " & strParm(.Fields("sSerialID")) & _
                     ", " & strParm(.Fields("sRemarksx")) & _
                     ", " & strParm(.Fields("sApproved")) & _
                     ", " & strParm(.Fields("cTranStat")) & _
                     ", " & strParm(p_oAppDrivr.UserID) & _
                     ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      Else
         If UndoTransaction = False Then
            p_oAppDrivr.RollbackTrans
            GoTo endProc
         End If
         
         lsSQL = "UPDATE MC_Warranty SET" & _
                     ", dTransact = " & dateParm(.Fields("dTransact")) & _
                     ", sClientID = " & strParm(.Fields("sClientID")) & _
                     ", sSerialID = " & strParm(.Fields("sSerialID")) & _
                     ", sRemarksx = " & strParm(.Fields("sRemarksx")) & _
                     ", sApproved = " & strParm(lsApproved) & _
                     ", cTranStat = " & strParm(.Fields("cTranStat")) & _
                     ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                  " WHERE sTransNox = " & strParm(.Fields("sTransNox"))
      End If
      If p_oAppDrivr.Execute(lsSQL, "MC_Warranty") = 0 Then
         p_oAppDrivr.RollbackTrans
         MsgBox "Unable to Save MC Warranty Pullout!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   Set loMCTrans = New clsMotorcycle
   With loMCTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      
      .InitTransaction
      .TransactDate = p_oRecMaster("dTransact")
      .MCInventoryID = p_oRecMaster("sMCInvIDx")
      .Quantity = 1
      
      .Serial(0, "SerialID") = p_oRecMaster("sSerialID")
      .Serial(0, "EngineNo") = p_oRecMaster("sEngineNo")
      If .WarrantyPullOut(p_oRecMaster("sTransNox")) = False Then
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
   End With
   p_oAppDrivr.CommitTrans
   
   p_nEditMode = xeModeReady
   pbRecExist = True
   SaveTransaction = True

endProc:
   Set loMCTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function PostTransaction(sTransNo As String, dTransact As Date) As Boolean
   Dim loMCTrans As clsMotorcycle
   Dim lsOldProc As String, lsSQL As String
   Dim lnRow As Long

   lsOldProc = "PostTransaction"
   'On Error Goto errProc
   PostTransaction = False
   
   If pbInitTran = False Or _
         (pbRecExist = False And p_nEditMode = xeModeReady) Then GoTo endProc

   p_oAppDrivr.BeginTrans
   With p_oRecMaster
      lsSQL = "UPDATE MC_Warranty SET" & _
                  ", cTranStat = " & strParm(xeStatePosted) & _
                  ", dReleased = " & dateParm(dTransact) & _
                  ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(.Fields("sTransNox"))
      
      If p_oAppDrivr.Execute(lsSQL, "MC_Warranty") = 0 Then
         p_oAppDrivr.RollbackTrans
         MsgBox "Unable to Release Warranty!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   Set loMCTrans = New clsMotorcycle
   With loMCTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      
      .InitTransaction
      .TransactDate = p_oRecMaster("dTransact")
      .MCInventoryID = p_oRecMaster("sMCInvIDx")
      .Quantity = 1
      
      .Serial(0, "SerialID") = p_oRecMaster("sSerialID")
      .Serial(0, "EngineNo") = p_oRecMaster("sEngineNo")
      If .WarrantyRelease(p_oRecMaster("sTransNox")) = False Then
         p_oAppDrivr.RollbackTrans
         GoTo endProc
      End If
      
      p_oAppDrivr.CommitTrans
   End With
   
   p_nEditMode = xeModeReady
   pbRecExist = True
   PostTransaction = True

endProc:
   Set loMCTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & sTransNo _
                       & ", " & dTransact & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsOldProc As String, lsSQL As String
   Dim lnRow As Long

   lsOldProc = "OpenTransaction"
   'On Error Goto errProc
   OpenTransaction = False

   If pbInitTran = False Then GoTo endProc

   lsSQL = AddCondition(p_sActQuery, "a.sTransNox = " & strParm(sTransNo))
   
   Set p_oRecMaster = New Recordset
   With p_oRecMaster
      .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
      
      If .EOF Then
         InitTransaction
         GoTo endProc
      End If
      
      p_sSerialID = .Fields("sSerialID")
      p_sFullName = .Fields("xFullName")
      p_sAddressx = .Fields("xAddressx")
      p_cTranStat = .Fields("cTranStat")
   End With
   
   p_nEditMode = xeModeReady
   pbRecExist = True
   OpenTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbRecExist = False Then Exit Function
   
   If p_nEditMode = xeModeReady Then Exit Function
   
   p_nEditMode = xeModeUpdate
   UpdateTransaction = True
End Function

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = "DeleteTransaction"
   'On Error Goto errProc
   DeleteTransaction = False

   If pbInitTran = False Or pbRecExist = False Then GoTo endProc

   p_nEditMode = xeModeDelete
   
   With p_oAppDrivr
      .BeginTrans
      lsSQL = "DELETE FROM MC_Warranty" & _
               " WHERE sTransNox = " & strParm(p_oRecMaster("sTransNox"))
      If .Execute(lsSQL, "MC_Warranty") = 0 Then
         .RollbackTrans
         GoTo endProc
      End If
      
      If UndoTransaction = False Then
         .RollbackTrans
         GoTo endProc
      End If
      
      .CommitTrans
   End With
   p_nEditMode = xeModeReady
   pbRecExist = False
   DeleteTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsOldProc = "SearchTransaction"
   'On Error Goto errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   If Not IsMissing(sSearch) Then
      lsCondition = "a.sTransNox = " & strParm(Trim(sSearch))
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lasSelect = GetSplitedName(sSearch)
            lsCondition = "b.sLastName LIKE " & strParm(lasSelect(0) & "%") & _
                           " AND (b.sFrstName LIKE " & strParm(lasSelect(1) & "%") & _
                              " OR b.sFrstName LIKE " & strParm(lasSelect(1) & lasSelect(2) & "%") & _
                              IIf(lasSelect(2) = Empty, " )", _
                                 " OR b.sMiddName LIKE " & strParm(lasSelect(2) & "%") & ")")
         End If
      End If
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.dTransact" & _
                  ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
                  ", c.sEngineNo" & _
               " FROM MC_Warranty a" & _
                  ", Client_Master b" & _
                  ", MC_Serial c" & _
               " WHERE a.sClientID = b.sClientID" & _
                  " AND a.sSerialID = c.sSerialID"
                  
      If lsCondition <> Empty Then lsSQL = lsSQL & " AND " & lsCondition
      
      lsSQL = lsSQL & " ORDER BY xFullName"
      
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then
         If Not IsMissing(sSearch) Then InitTransaction
         GoTo endProc
      End If

      If lors.RecordCount = 1 Then
         SearchTransaction = OpenTransaction(lors("sTransNox"))
         GoTo endProc
      End If

      lsSQL = KwikBrowse(p_oAppDrivr, lors, "sTransNox»dTransact»xFullName»sEngineNo", _
                  "Trans No»Date»Customer Name»EngineNo", "@@@@-@@@@@@»MM/DD/YYYY»@»@")
                  
      If lsSQL = Empty Then
         If Not IsMissing(sSearch) Then InitTransaction
         GoTo endProc
      End If
      lasSelect = Split(lsSQL, "»")
      SearchTransaction = OpenTransaction(lasSelect(0))
   End With

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & sSearch _
                       & ", " & bByCode & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String, lsSQL As String
   Dim lnRow As Long

   lsOldProc = "CancelTransaction"
   'On Error Goto errProc
   CancelTransaction = False

   If pbInitTran = False Or pbRecExist = False Then GoTo endProc

   With p_oAppDrivr
      .BeginTrans
      
      lsSQL = "UPDATE MC_Warranty SET" & _
                  "  cTranStat = " & strParm(xeStateCancelled) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(p_oRecMaster("sTransNox"))
      If .Execute(lsSQL, "MC_Warranty") = 0 Then
         .RollbackTrans
         MsgBox "Unable to Update Cash Disbursement Voucher!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If UndoTransaction = False Then
         .RollbackTrans
         GoTo endProc
      End If
      
      .CommitTrans
   End With
   CancelTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function SearchMaster(ByVal Index As Integer, Optional lsSearch As Variant) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SearchMaster"
   'On Error Goto errProc
   
   If Index = 2 Then
      If Not IsMissing(lsSearch) Then
         SearchMaster = getClient(lsSearch)
      Else
         SearchMaster = getClient
      End If
   End If

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & Index _
                        & ", " & lsSearch & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, "clsMCWarranty", lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Function getClient(Optional lsSearch As Variant) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsCondition As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsOldProc = "GetClient"
   'On Error Goto errProc
   getClient = False

   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(lsSearch) Then
      If Trim(lsSearch) = "" Then GoTo endWithClear
      If lsSearch = p_sFullName Then GoTo endProc
      
      lasSelect = GetSplitedName(lsSearch)
      lsCondition = " a.sLastName LIKE " & strParm(lasSelect(0) & "%") & _
                     " AND (a.sFrstName LIKE " & strParm(lasSelect(1) & "%") & _
                        " OR a.sFrstName LIKE " & strParm(lasSelect(1) & lasSelect(2) & "%") & _
                        IIf(lasSelect(2) = Empty, " )", _
                           " OR a.sMiddName LIKE " & strParm(lasSelect(2) & "%") & ")")
   End If

   With p_oAppDrivr
      lsSQL = "SELECT" & _
                  "  a.sClientID" & _
                  ", CONCAT(a.sLastName, ', ', a.sFrstName) AS xFullName" & _
                  ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                  ", d.sEngineNo" & _
                  ", d.sFrameNox" & _
                  ", e.sModelNme" & _
                  ", d.sSerialID" & _
                  ", d.sMCInvIDx" & _
                  ", f.dTransact" & _
               " FROM Client_Master a" & _
                  ", TownCity b" & _
                  ", Province c" & _
                  ", MC_Serial d" & _
                  ", MC_Model e" & _
                  ", MC_SO_Master f" & _
                  ", MC_SO_Detail g" & _
               " WHERE a.sTownIDxx = b.sTownIDxx" & _
                  " AND b.sProvIDxx = c.sProvIDxx" & _
                  " AND a.sClientID = d.sClientID" & _
                  " AND d.sModelIDx = e.sModelIDx" & _
                  " AND f.sTransNox = g.sTransNox" & _
                  " AND a.sClientID = f.sClientID" & _
                  " AND d.sSerialID = g.sSerialID"

      If lsCondition <> Empty Then lsSQL = lsSQL & " AND " & lsCondition
      Set lors = New Recordset
      lors.Open lsSQL, .Connection, , , adCmdText

      If lors.EOF Then
         If Not IsMissing(lsSearch) Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If

      With p_oRecMaster
         If lors.RecordCount = 1 Then
            .Fields("sClientID") = lors("sClientID")
            .Fields("sSerialID") = lors("sSerialID")
            .Fields("sEngineNo") = lors("sEngineNo")
            .Fields("sFrameNox") = lors("sFrameNox")
            .Fields("sModelNme") = lors("sModelNme")
            .Fields("sMCInvIDx") = lors("sMCInvIDx")
            .Fields("dPurchase") = lors("dTransact")
            p_sFullName = lors("xFullName")
            p_sAddressx = lors("xAddressx")
            RaiseEvent MasterRetrieved(2)
            RaiseEvent MasterRetrieved(3)
            RaiseEvent MasterRetrieved(4)
            RaiseEvent MasterRetrieved(5)
            RaiseEvent MasterRetrieved(6)
            RaiseEvent MasterRetrieved(7)
         Else
            lsSQL = KwikBrowse(p_oAppDrivr, lors, "sClientID»xFullName»xAddressx", "Client ID»Customer Name»Address")
            If lsSQL = Empty Then
               If Not IsMissing(lsSearch) Then GoTo endWithClear
            Else
               .Fields("sClientID") = lors("sClientID")
               .Fields("sSerialID") = lors("sSerialID")
               .Fields("sEngineNo") = lors("sEngineNo")
               .Fields("sFrameNox") = lors("sFrameNox")
               .Fields("sModelNme") = lors("sModelNme")
               .Fields("dPurchase") = lors("dTransact")
               p_sFullName = lors("xFullName")
               p_sAddressx = lors("xAddressx")
               RaiseEvent MasterRetrieved(2)
               RaiseEvent MasterRetrieved(3)
               RaiseEvent MasterRetrieved(4)
               RaiseEvent MasterRetrieved(5)
               RaiseEvent MasterRetrieved(6)
               RaiseEvent MasterRetrieved(7)
            End If
         End If
      End With
   End With
   getClient = True

endProc:
   Set lors = Nothing

   Exit Function
endWithClear:
   p_sFullName = ""
   p_sAddressx = ""
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   RaiseEvent MasterRetrieved(5)
   RaiseEvent MasterRetrieved(7)
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & lsSearch & " )"
   GoTo endProc
End Function

Private Function UndoTransaction() As Boolean
   Dim loMCTrans As clsMotorcycle
   Dim lsOldProc As String
   
   lsOldProc = "UndoTransaction"
   'On Error Goto errProc

   Set loMCTrans = New clsMotorcycle
   With loMCTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      
      .InitTransaction
      .TransactDate = p_oRecMaster("dTransact")
      .MCInventoryID = p_oRecMaster("sMCInvIDx")
      .Quantity = 1
      
      .Serial(0, "SerialID") = p_oRecMaster("sSerialID")
      .Serial(0, "EngineNo") = p_oRecMaster("sEngineNo")
      If .WarrantyPullOut(p_oRecMaster("sTransNox"), True) = False Then GoTo endProc
   End With
   
   UndoTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function HasRights2Save(ByRef lsUserID As String) As Boolean
   Dim lsOldProc As String
   Dim lnUserRght As Integer, lsUserName As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   
   lsOldProc = "HasRights2Save"
   'On Error Goto errProc
   HasRights2Save = False
   
   lsUserID = p_oAppDrivr.UserID
   If p_oRecMaster("cTranStat") = xeStatePosted Or p_oRecMaster("cTranStat") = xeStateCancelled Then
      If p_oAppDrivr.UserLevel = xeSysMaster Then
         HasRights2Save = True
      Else
         MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
                  "Request not granted!!!", vbCritical, "Warning"
      End If
      GoTo endProc
   End If
   
   With p_oAppDrivr
      If .MDIMain.Controls(.MenuName).Tag = "" Then
         ReDim lasRights(4) As String
   
         lasRights(1) = 192
         lasRights(2) = 192
         lasRights(3) = 192
      Else
         If InStr(0, .MDIMain.Controls(.MenuName).Tag, "»", vbTextCompare) = 0 Then
            ReDim lasRights(4) As String
      
            lasRights(1) = 192
            lasRights(2) = 192
            lasRights(3) = 192
         Else
            lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
         End If
      End If
   
      lbSeek = False
      If p_nEditMode = xeModeAddNew Then
         ' Add Rights
         lnRights = CInt(lasRights(1))
         If lnRights And .UserLevel = 0 Then
            lnSeek = MsgBox("User Doesn't Have the Right to Create New Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
            If lnSeek = vbNo Then GoTo endProc
            lbSeek = True
         End If
      ElseIf p_nEditMode = xeModeUpdate Then
         ' update Rights
         lnRights = CInt(lasRights(2))
   
         If lnRights And .UserLevel = 0 Then
            Select Case .UserLevel
            Case xeEncoder
               lbSeek = True
            Case xeSupervisor
               If DateDiff("d", p_oRecMaster("dModified"), p_oAppDrivr.ServerDate) > 2 Then lbSeek = True
            Case xeManager
               If DateDiff("d", p_oRecMaster("dModified"), p_oAppDrivr.ServerDate) > 30 Then lbSeek = True
            Case xeAudit
               lbSeek = True
            Case xeSysAdmin
               lbSeek = True
            End Select
         End If
      End If
   
      If lbSeek = False Then
         HasRights2Save = True
         GoTo endProc
      End If
   
      lnSeek = MsgBox("User Doesn't Have the Right to Update Transaction!!!" & _
                  vbCrLf & "Seek Authorized User's Approval?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc
   
      If GetApproval(p_oAppDrivr, lnUserRght, lsUserID, lsUserName, p_oAppDrivr.MenuName) = False Then GoTo endProc
   
      If p_nEditMode = xeModeAddNew Then
         If lnUserRght And .UserLevel = 0 Then
            MsgBox "Approving User Doesn't Have the Right to Create New Transaction!!!" & _
            "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      ElseIf p_nEditMode = xeModeUpdate Then
         If lnUserRght And .UserLevel = 0 Then
            Select Case lnUserRght
            Case xeSupervisor
               If DateDiff("d", p_oRecMaster("dModified"), p_oAppDrivr.ServerDate) <= 2 Then lbSeek = False
            Case xeManager
               If DateDiff("d", p_oRecMaster("dModified"), p_oAppDrivr.ServerDate) <= 31 Then lbSeek = False
            Case xeSysAdmin
               If .UserLevel <> xeSysAdmin Then lbSeek = False
            End Select
      
            If lbSeek Then
               MsgBox "Approving User Doesn't Have the Right to Update Transaction!!!" & _
               "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
   End With
   HasRights2Save = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsUserID & " )"
End Function
