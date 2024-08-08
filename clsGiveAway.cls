VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGiveAway"
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
'     Motorcycle Sales GiveAways Class Object
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
'  XerSys [ 12/09/2004 01:29 pm ]
'     Start creating this object.
'  XerSys [ 12/10/2004 10:40 am ]
'     Instead of creating a class independent on master-detail base class, use the later
'        as its base class
'  XerSys [ 02/21/2005 05:00 pm ]
'     Adjust this object to fit to the new implementation of the application driver
'  Kalyptus [ 12/11/2006 03:28 pm]
'     Implements the BugTracker
'  XerSys [ 09/28/2007 09:33 am ]
'     Update this object to adopt the new implementation of SP Inventory Transaction.
'  Jheff [ 01/18/2008 11:05 am ]
'     Update this object for interface transaction
'  XerSys [ 02/06/2008 09:19 am ]
'     Include Giveaway swapping for direct usage of interface transaction
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsGiveAway"
'Private Enum pxtGiveAway
Private Const pxeOriginal = 0
Private Const pxeReplaced As Integer = 1
Private Const pxeRemoved As Integer = 2
Private Const pxeAdded = 3
'End Enum

Private Const pxeMCSales = "MCSl"

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private WithEvents p_oSPSales As clsSPPOSBranch
Attribute p_oSPSales.VB_VarHelpID = -1
Private p_oModel As Recordset

Private p_sBranchCd As String
Private p_sParentxx As String
Private p_nReplAmtx As Double

Private p_sFullName As String
Private p_sAddressx As String
Private p_cPaymForm As String

Private pbInitTran As Boolean
Private pbShowMsg As Boolean
Private pbHasSPRep As Boolean
Private pbSPRepOkx As Boolean
Private pbReleased As Boolean
Private p_bIsDelete As Boolean

Private pnCtr As Integer

Public Event DetailRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let Released(ByVal lbReleased As Boolean)
   pbReleased = lbReleased
End Property

Property Get Detail(ByVal Row As Long, Index As Variant) As Variant
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   Detail = p_oBaseClas.Detail(Row, Index)
End Property

Property Let Detail(ByVal Row As Long, Index As Variant, Value As Variant)
   On Error Resume Next
   
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 1, "sbarrcode"
      getDetail Row, Index, Value, False
   Case 2, "sdescript"
      getDetail Row, Index, Value, False
   Case 7, "cgawystat"
      Select Case Value
      Case pxeOriginal, pxeReplaced, pxeRemoved, pxeAdded
         p_oBaseClas.Detail(Row, Index) = Value
      Case Else
         p_oBaseClas.Detail(Row, Index) = pxeOriginal
      End Select
      Call computeReplacement
   Case 6, "ngivenxxx"
      p_oBaseClas.Detail(Row, Index) = Value
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   'mac 2022.10.22
   '  pass payment form for checking mc sales promo giveaways
   Case 80, "cpaymform"
      Master = p_cPaymForm
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   ' enable this if the end of day transaction is okie
   'Case 1, "dtransact"
      'p_obaseclas.Master(Index) = getValidTranDate(p_oAppDrivr.SysDate, Value)
   Case 3, "xfullname"
      p_sFullName = Value
   Case 4, "xaddressx"
      p_sAddressx = Value
   'mac 2022.10.22
   '  pass payment form for checking mc sales promo giveaways
   Case 80, "cpaymform"
      p_cPaymForm = Value
   Case Else
      p_oBaseClas.Master(Index) = Value
   End Select
End Property

Property Let Parent(ByVal Value As String)
   p_sParentxx = Value
End Property

Property Get ReplaceAmount() As Double
   ReplaceAmount = p_nReplAmtx
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
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
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "MC_SO_Master"
      .DetailTable = "MC_SO_GiveAways"
      .DisplayMessage = False
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDRNoxxxx" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) xFullName" & _
                        ", CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) xAddressx" & _
                        ", a.nTranTotl" & _
                        ", a.nAmtPaidx" & _
                        ", a.sORNoxxxx" & _
                        ", a.sApproved" & _
                        ", a.sRemarksx" & _
                        ", a.cPaymForm" & _
                        ", a.cTranType" & _
                        ", a.cCustType" & _
                        ", a.cTranStat" & _
                        ", a.sApplicNo" & _
                        ", a.sClientID" & _
                        ", a.sModified" & _
                        ", a.dModified"
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx"
  
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sBarrCode" & _
                        ", b.sDescript" & _
                        ", c.nQtyOnHnd" & _
                        ", b.nSelPrice" & _
                        ", a.nQuantity" & _
                        ", a.nGivenxxx" & _
                        ", a.cGAwyStat" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
                        ", a.sPartsIDx" & _
                        ", c.nResvOrdr" & _
                        ", c.nBackOrdr" & _
                        ", c.nReorderx" & _
                        ", c.nLedgerNo" & _
                     " FROM " & .DetailTable & " a" & _
                        ", Spareparts b" & _
                        ", SP_Inventory c"
      .DetailQuery = .DetailQuery & _
                     " WHERE a.sPartsIDx = b.sPartsIDx" & _
                        " AND a.sPartsIDx = c.sPartsIDx" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
         
      If pbReleased Then
         .DetailQuery = AddCondition(.DetailQuery, " ( a.nQuantity > a.nGivenxxx" & _
                        " AND ( a.cGAwyStat = '0' OR a.cGAwyStat = '3' ) )")
      End If
      
      .Reference = "a.sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sDRNoxxxx" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                           ", MC_SO_GiveAways c" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
                           " AND a.sTransNox = c.sTransNox" & _
                           " AND a.cTranStat <> " & strParm(xeStateCancelled)
      If pbReleased Then
         .BrowseQuery(0) = AddCondition(.BrowseQuery(0), " ( c.nQuantity > c.nGivenxxx" & _
                        " AND ( c.cGAwyStat = '0' OR c.cGAwyStat = '3' ) )")
      End If
      
      .BrowseQuery(0) = .BrowseQuery(0) & _
                        " GROUP BY a.sDRNoxxxx" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName)" & _
                        " ORDER BY a.sTransNox" & _
                           ", a.dTransact" & _
                           ", xFullName"
                           
      .BrowseColumn(0) = "sDRNoxxxx»xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sDRNoxxxx»CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "DR No»Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@»@»@»MMMM DD, YYYY"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sPartsIDx" & _
                                 ", b.sBarrCode" & _
                                 ", b.sDescript" & _
                                 ", a.nQtyOnHnd" & _
                                 ", b.nSelPrice" & _
                                 ", a.nResvOrdr" & _
                                 ", a.nBackOrdr" & _
                                 ", a.nReorderx" & _
                                 ", a.nLedgerNo" & _
                              " FROM SP_Inventory a" & _
                                 ", Spareparts b" & _
                              " WHERE a.sPartsIDx = b.sPartsIDx" & _
                                 " AND a.sBranchCd = " & strParm(p_sBranchCd)
      .BrowseDetailColumn(0) = "sBarrCode»sDescript»nQtyOnHnd"
      .BrowseDetailTitle(0) = "Bar Code»Description»QOH"
   End With
   
   Set p_oSPSales = New clsSPPOSBranch
   With p_oSPSales
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc
   End With
   
   pbInitTran = InitTransaction
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Sub InitGiveaway()
   Set p_oModel = New Recordset
   With p_oModel
      .Fields.Append "sModelIDx", adVarChar, Len(p_sBranchCd) + 8
      .Fields.Append "nQuantity", adInteger
      .Open
   End With
   
   Call p_oBaseClas.clearDetail
   pbHasSPRep = False
End Sub

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   'On Error GoTo errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsAppvID As String, lsAppvName As String
   Dim lnAppvRights As Integer, lnAllowRights As Integer
   Dim lbGetApproval As Boolean
   Dim lnCtr As Integer, lnEntryNo As Integer
   
   lsProcName = "p_oBaseClas_WillSave"
   'On Error GoTo errProc
   
   If pbShowMsg Then
      If MsgBox("Are you sure you want to Save this Transaction?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
   End If
   
   With p_oBaseClas
      If .isDetailModify = False Then
         'Added testing of changes in dTransact
         If .Master("dTransact") = .MasterOriginal("dTransact") Then
            SaveTransaction = True
            GoTo endProc
         End If
      End If
      
      If p_nReplAmtx > 0 And p_oSPSales.Master("nReplAmtx") <> p_nReplAmtx Then
         If GetSPReplacement() = False Then GoTo endProc
      End If
         
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "cGAwyStat") = pxeAdded Then
            lbGetApproval = True
            Exit For
         End If
      Next
      
      If lbGetApproval Then
         lnAllowRights = 255 'p_oAppDrivr.getConfiguration("MCGAway")
         If lnAllowRights And p_oAppDrivr.UserLevel = 0 Then
            lbGetApproval = MsgBox("User Doesn't Have the Right to Add New Giveaway!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbYes

            If GetApproval(p_oAppDrivr, _
                  lnAppvRights, _
                  lsAppvID, _
                  lsAppvName, _
                  p_oAppDrivr.MenuName) = False Then GoTo endProc

            If (lnAllowRights And lnAppvRights) = 0 Then
               MsgBox "Approving Officer Has no Right to Add New Giveaway!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
      
      Set loSPTrans = New clsSPInventoryTrans
      With loSPTrans
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction() = False Then GoTo endProc
      End With
      
      If .EditMode = xeModeUpdate Then
         lsSQL = "SELECT nEntryNox" & _
                  " FROM " & .DetailTable & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                  " ORDER BY nEntryNox DESC" & _
                  " LIMIT 1"
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            
         If loRS.EOF Then
            lnEntryNo = 0
         Else
            lnEntryNo = loRS("nEntryNox")
         End If
      End If
      
      lnCtr = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "nEntryNox") = 0 Then
            '.Detail(pnCtr, "nEntryNox") = lnEntryNo
            lnEntryNo = lnEntryNo + 1
         End If
         
         ' filter out remove and replace giveaway
         If .Detail(pnCtr, "cGAwyStat") = pxeAdded Or _
               .Detail(pnCtr, "cGAwyStat") = pxeOriginal Then
            If .Detail(pnCtr, "nGivenxxx") > 0 Then
               
               '2009.07.18 - Kalyptus
               'Check the Quantity on Hand
               If .Detail(pnCtr, "nGivenxxx") > Detail(pnCtr, "nQtyOnHnd") Then
                    MsgBox "There are items that cannot be release as MC Sales GiveAways!!!" & _
                    vbCrLf & " Please check your entries...", vbCritical, "Warning"
                    GoTo endProc
               End If
               
               loSPTrans.Detail(lnCtr, "sPartsIDx") = .Detail(pnCtr, "sPartsIDx")
               loSPTrans.Detail(lnCtr, "nQuantity") = .Detail(pnCtr, "nGivenxxx")
               loSPTrans.Detail(lnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd")
               loSPTrans.Detail(lnCtr, "nResvOrdr") = .Detail(pnCtr, "nResvOrdr")
               loSPTrans.Detail(lnCtr, "nBackOrdr") = .Detail(pnCtr, "nBackOrdr")
               loSPTrans.Detail(lnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
               
               lnCtr = lnCtr + 1
               
            End If
         End If
            
         If .EditMode = xeModeAddNew Then
            lsSQL = "INSERT INTO MC_SO_GiveAways SET" & _
                        "  sTransNox = " & strParm(.Master("sTransNox")) & _
                        ", nEntryNox = " & pnCtr + 1 & _
                        ", sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                        ", nQuantity = " & .Detail(pnCtr, "nQuantity") & _
                        ", nGivenxxx = " & .Detail(pnCtr, "nGivenxxx") & _
                        ", cGAwyStat = " & strParm(.Detail(pnCtr, "cGAwyStat")) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
         Else
            lsSQL = "SELECT sPartsIDx" & _
                     " FROM MC_SO_GiveAways" & _
                     " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                        " AND nEntryNox = " & pnCtr + 1

            Set loRS = New Recordset
            loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

            If .Detail(pnCtr, "nEntryNox") = 0 Then
               lsSQL = "INSERT INTO MC_SO_GiveAways SET" & _
                           "  sTransNox = " & strParm(.Master("sTransNox")) & _
                           ", nEntryNox = " & lnEntryNo & _
                           ", sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                           ", nQuantity = " & .Detail(pnCtr, "nQuantity") & _
                           ", nGivenxxx = " & .Detail(pnCtr, "nGivenxxx") & _
                           ", cGAwyStat = " & strParm(.Detail(pnCtr, "cGAwyStat")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate())
            Else
               lsSQL = "UPDATE MC_SO_GiveAways SET" & _
                           "  sPartsIDx = " & strParm(.Detail(pnCtr, "sPartsIDx")) & _
                           ", nQuantity = " & strParm(.Detail(pnCtr, "nQuantity")) & _
                           ", nGivenxxx = " & strParm(.Detail(pnCtr, "nGivenxxx")) & _
                           ", cGAwyStat = " & strParm(.Detail(pnCtr, "cGAwyStat")) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                        " WHERE sTransNox = " & strParm(.Master("sTransNox")) & _
                        " AND nEntryNox = " & .Detail(pnCtr, "nEntryNox")
            End If
         End If

         If p_oAppDrivr.Execute(lsSQL, "MC_SO_GiveAways", p_sBranchCd) = 0 Then
            MsgBox "Unable to Update MC Sales GiveAways!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Next
      
      ' save SP Transaction if a valid record exists
      If lnCtr > 0 Then
         If loSPTrans.SalesGiveAway(.Master("sTransNox"), _
               .Master("dTransact"), _
               IIf(p_sParentxx = "clsGiveAway", .EditMode, xeModeAddNew)) = False Then
            GoTo endProc
         End If
      End If
   End With

   With p_oSPSales
      If .Detail(0, "sPartsIDx") <> Empty Then
         .Master("dTransact") = p_oBaseClas.Master("dTransact")
         .Master("sClientID") = p_oBaseClas.Master("sClientID")
         .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
         .Master("sSourceCd") = pxeMCSales
         If .SaveTransaction() = False Then GoTo endProc

         If loSPTrans.InitTransaction() = False Then GoTo endProc
      ElseIf pbHasSPRep And .Detail(0, "sPartsIDx") = Empty Then
         If .DeleteTransaction = False Then GoTo endProc
      End If
   End With
      
   SaveTransaction = True
   
endProc:
   Set loSPTrans = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   CancelTransaction = delSPTransaction()

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   'On Error GoTo errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "DeleteTransaction"
   'On Error GoTo errProc
   
   If delSPTransaction() = False Then GoTo endProc

   With p_oBaseClas
      lsSQL = "DELETE FROM " & .DetailTable & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
               
      If p_oAppDrivr.Execute(lsSQL, .DetailTable, p_sBranchCd) = 0 Then
         MsgBox "Unable to Delete MC Sales GiveAways!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If pbHasSPRep Then
         If p_oSPSales.DeleteTransaction() = False Then GoTo endProc
      End If
   End With
   
   DeleteTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, _
                           Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsProcName = "SearchTransaction"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sDRNoxxxx"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lasMaster = GetSplitedName(sSearch)
            lsValue = Trim(lasMaster(0)) & "»" & Trim(lasMaster(1))
            lsField = "b.sLastName" & "»" & "b.sFrstName"
            lsDescript = "b.sLastName" & "»" & "b.sFrstName"
         End If
      End If
   End If
   
   With p_oBaseClas
      If .SearchTransaction(lsValue, lsField, lsDescript) Then
         If .Detail(0, "sTransNox") = Empty Then
            .Detail(0, "nEntryNox") = 1
            .Detail(0, "sDescript") = ""
            .Detail(0, "nQtyOnHnd") = 0
            .Detail(0, "nQuantity") = 0
            .Detail(0, "nGivenxxx") = 0
            .Detail(0, "sTransNox") = .Master("sTransNox")
            .Detail(0, "sPartsIDx") = ""
         End If
         SearchTransaction = True
      End If
   End With

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function AddDetail() As Boolean
   With p_oBaseClas
      AddDetail = .AddDetail
      
      pnCtr = .ItemCount - 1
      If .EditMode = xeModeAddNew Then
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      Else
         .Detail(pnCtr, "nEntryNox") = 0
      End If
      .Detail(pnCtr, "sDescript") = ""
      .Detail(pnCtr, "nQtyOnHnd") = 0
      .Detail(pnCtr, "nQuantity") = 0
      .Detail(pnCtr, "nGivenxxx") = 0
      .Detail(pnCtr, "cGAwyStat") = pxeAdded
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "sPartsIDx") = ""
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Sub AddModel(Value As String)
   
   If Left(p_sParentxx, 3) <> "cls" Then Exit Sub
   
   With p_oModel
      Call .Find("sModelIDx = " & strParm(Value), 1, adSearchForward)
      
      If .EOF Then
         .AddNew
         
         .Fields("sModelIDx") = Value
         .Fields("nQuantity") = 1
      Else
         .Fields("nQuantity") = .Fields("nQuantity") + 1
      End If
   End With
End Sub

Function SearchDetail(ByVal Row As Long, _
                      ByVal Index As Variant, _
                      Optional Value As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, 2, "sbarrcode", "sdescript"
      SearchDetail = getDetail(Row, Index, Value, True)
   End Select
'   SearchDetail = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
End Function

Function LoadDetail() As Boolean
   Dim loRS As Recordset, loGiveAway As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lbGiven As Boolean
   
   lsProcName = "LoadDetail"
   'On Error GoTo errProc
   
   If p_oModel.RecordCount = 0 Then GoTo endProc
   
   Set loGiveAway = New Recordset
   With loGiveAway
      .Fields.Append "sPartsIDx", adVarChar, Len(p_sBranchCd) + 8
      .Fields.Append "sBarrCode", adVarChar, 25
      .Fields.Append "nQuantity", adInteger
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "nQtyOnHnd", adInteger
      .Fields.Append "nSelPrice", adInteger
      .Fields.Append "nResvOrdr", adInteger
      .Fields.Append "nBackOrdr", adInteger
'190         .Fields.Append "nReorderx", adInteger
      .Fields.Append "nLedgerNo", adInteger
      .Open
   End With
   
   With p_oModel
      'this one added if .recordcount > 0 then
      If .RecordCount > 0 Then .MoveFirst
      Do While .EOF = False
         lsSQL = "SELECT" & _
                     "  a.sPartsIDx" & _
                     ", b.sDescript" & _
                     ", a.nQuantity" & _
                     ", b.nSelPrice" & _
                     ", c.nQtyOnHnd" & _
                     ", c.nResvOrdr" & _
                     ", c.nBackOrdr" & _
                     ", c.nReorderx" & _
                     ", c.nLedgerNo" & _
                     ", b.sBarrCode" & _
                  " FROM MC_Sales_GiveAways a" & _
                     ", Spareparts b" & _
                     ", SP_Inventory c" & _
                  " WHERE a.sPartsIDx = b.sPartsIDx" & _
                     " AND b.sPartsIDx = c.sPartsIDx" & _
                     " AND c.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " AND a.sModelIDx = " & strParm(.Fields("sModelIDx"))
         Debug.Print lsSQL
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         Do While loRS.EOF = False
            loGiveAway.Find "sPartsIDx = " & strParm(loRS("sPartsIDx")), 0, adSearchForward, 1
            
            If loGiveAway.EOF Then
               loGiveAway.AddNew
               loGiveAway("sPartsIDx") = loRS("sPartsIDx")
               loGiveAway("sDescript") = loRS("sDescript")
               loGiveAway("nQuantity") = loRS("nQuantity") * .Fields("nQuantity")
               loGiveAway("nQtyOnHnd") = loRS("nQtyOnHnd")
               loGiveAway("nSelPrice") = loRS("nSelPrice")
               loGiveAway("nResvOrdr") = loRS("nResvOrdr")
               loGiveAway("nBackOrdr") = loRS("nBackOrdr")
               loGiveAway("nLedgerNo") = loRS("nLedgerNo")
               loGiveAway("sBarrCode") = loRS("sBarrCode")
            Else
               loGiveAway("nQuantity") = loGiveAway("nQuantity") + _
                                          (loRS("nQuantity") * .Fields("nQuantity"))
            End If
            
            loRS.MoveNext
         Loop
         
         'kalyptus - 2022.10.18 01:39pm
         'Call the method to add the promotional giveaway
         Call loadPromoGiveaway(loGiveAway, .Fields("sModelIDx"), .Fields("nQuantity"))
         
         .MoveNext
      Loop
   End With
   
   With loGiveAway
      If .EOF = False Then .MoveFirst
      For pnCtr = 0 To .RecordCount - 1
         If pnCtr > p_oBaseClas.ItemCount - 1 Then AddDetail
         
         p_oBaseClas.Detail(pnCtr, "nEntryNox") = pnCtr + 1
         p_oBaseClas.Detail(pnCtr, "sPartsIDx") = .Fields("sPartsIDx")
         p_oBaseClas.Detail(pnCtr, "sDescript") = .Fields("sDescript")
         p_oBaseClas.Detail(pnCtr, "nQuantity") = .Fields("nQuantity")
         p_oBaseClas.Detail(pnCtr, "nGivenxxx") = .Fields("nQuantity")
         p_oBaseClas.Detail(pnCtr, "cGAwyStat") = pxeOriginal
         p_oBaseClas.Detail(pnCtr, "nQtyOnHnd") = .Fields("nQtyOnHnd")
         p_oBaseClas.Detail(pnCtr, "nSelPrice") = .Fields("nSelPrice")
         p_oBaseClas.Detail(pnCtr, "nResvOrdr") = .Fields("nResvOrdr")
         p_oBaseClas.Detail(pnCtr, "nBackOrdr") = .Fields("nBackOrdr")
         p_oBaseClas.Detail(pnCtr, "nLedgerNo") = .Fields("nLedgerNo")
         p_oBaseClas.Detail(pnCtr, "sBarrCode") = .Fields("sBarrCode")
         .MoveNext
      Next
   End With
   
   LoadDetail = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbInitTran = False Then Exit Function
   
   p_oBaseClas.EditMode = xeModeUpdate
   With p_oSPSales
      If .UpdateTransaction = False Then Exit Function
   End With
   
   UpdateTransaction = True
End Function

Function GetSPReplacement() As Boolean
   Dim loFormReplacement As frmGiveAwayReplace
   Dim lsProcName As String
   Dim lnReplAmtx As Double
   
   lsProcName = "getSPReplacement"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If p_nReplAmtx = 0 Then
      GetSPReplacement = True
      GoTo endProc
   End If
      
   Set loFormReplacement = New frmGiveAwayReplace
   With loFormReplacement
      ' assign first the replacement amount to a temporary variable
      lnReplAmtx = p_oSPSales.Master("nReplAmtx")
      
      p_oSPSales.Master("dTransact") = p_oBaseClas.Master("dTransact")
      p_oSPSales.Master("sClientID") = p_oBaseClas.Master("sClientID")
      p_oSPSales.Master("nReplAmtx") = p_nReplAmtx
      
      Set .AppDriver = p_oAppDrivr
      Set .SPSales = p_oSPSales
      .txtField(0) = Format(p_oSPSales.Master("sTransNox"), "@@@@-@@@@@@")
      .txtField(1) = Format(p_oSPSales.Master("dTransact"), "MMMM DD, YYYY")
      .txtField(2) = p_oSPSales.Master("sSalesInv")
      .txtField(3) = p_sFullName
      .txtField(4) = p_sAddressx
      .txtField(5) = Format(p_oSPSales.Master("nTranTotl"), "#,##0.00")
      .txtField(6) = Format(p_nReplAmtx, "#,##0.00")
      .txtField(7) = Format(p_oSPSales.Master("nTranTotl") - p_nReplAmtx, "#,##0.00")
      .txtField(8) = Format(p_oAppDrivr.UserName, ">")
      
      .Show 1
      
      If .Cancelled Then
         p_oSPSales.Master("nReplAmtx") = lnReplAmtx
         GoTo endProc
      End If
   End With
   
   GetSPReplacement = True
   
endProc:
   Unload loFormReplacement
   
   Exit Function
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

Private Function getDetail(ByVal lnRow As Integer, _
                      ByVal lnIndex As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "getDetail"
   'On Error GoTo errProc

   With p_oBaseClas
      If lsValue <> Empty Then
         If lnIndex = 1 Then     ' Bar Code
            If lsValue = .Detail(lnRow, "sBarrCode") Then GoTo endProc
            If lbSearch Then
               lsDetail = "b.sBarrCode LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sBarrcode = " & strParm(Trim(lsValue))
            End If
         Else                    ' Description
            If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
            If lbSearch Then
               lsDetail = "b.sDescript LIKE " & strParm(Trim(lsValue) & "%")
            Else
               lsDetail = "b.sDescript = " & strParm(Trim(lsValue))
            End If
         End If
      ElseIf lbSearch = False Then
         GoTo endWithClear
      End If

      lsDetail = .getDetail(lnRow, lsDetail)
      If lsDetail = Empty Then
         If lbSearch = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
   
      lasDetail = Split(lsDetail, "»")
      .Detail(lnRow, "sPartsIDx") = lasDetail(0)
      .Detail(lnRow, "sBarrCode") = lasDetail(1)
      .Detail(lnRow, "sDescript") = lasDetail(2)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(3)
      .Detail(lnRow, "nQuantity") = IIf(.Detail(lnRow, "nQtyOnHnd") > 0, 1, 0)
      .Detail(lnRow, "nGivenxxx") = 0
      .Detail(lnRow, "cGAwyStat") = pxeAdded
      .Detail(lnRow, "nSelPrice") = lasDetail(4)
      .Detail(lnRow, "nResvOrdr") = lasDetail(5)
      .Detail(lnRow, "nBackOrdr") = lasDetail(6)
      .Detail(lnRow, "nLedgerNo") = lasDetail(8)
   End With
   getDetail = True
   
endProc:
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(5)
   RaiseEvent DetailRetrieved(6)

   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sPartsIDx") = ""
      .Detail(lnRow, "sBarrCode") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
      .Detail(lnRow, "nQuantity") = 0
      .Detail(lnRow, "nGivenxxx") = 0
      .Detail(lnRow, "cGAwyStat") = pxeOriginal
      .Detail(lnRow, "nSelPrice") = 0
      .Detail(lnRow, "nResvOrdr") = 0
      .Detail(lnRow, "nBackOrdr") = 0
      .Detail(lnRow, "nLedgerNo") = 0
   End With
   getDetail = False
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function delSPTransaction() As Boolean
   Dim loSPTrans As clsSPInventoryTrans
   Dim lsProcName As String
   
   lsProcName = "delSPTransaction"
   'On Error GoTo errProc
   
   Set loSPTrans = New clsSPInventoryTrans
   With loSPTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction() = False Then GoTo endProc
      
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "sPartsIDx") = p_oBaseClas.Detail(pnCtr, "sPartsIDx")
         .Detail(pnCtr, "nQuantity") = p_oBaseClas.Detail(pnCtr, "nQuantity")
         .Detail(pnCtr, "nQtyOnHnd") = p_oBaseClas.Detail(pnCtr, "nQtyOnHnd")
         .Detail(pnCtr, "nResvOrdr") = p_oBaseClas.Detail(pnCtr, "nResvOrdr")
         .Detail(pnCtr, "nBackOrdr") = p_oBaseClas.Detail(pnCtr, "nBackOrdr")
         .Detail(pnCtr, "nLedgerNo") = p_oBaseClas.Detail(pnCtr, "nLedgerNo")
      Next
      
      ' if a serial is not successfully updated, cancel saving.
      If .SalesGiveAway(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("dTransact"), _
            xeModeDelete) = False Then
         GoTo endProc
      End If
   End With
   
   delSPTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub computeReplacement()
   With p_oBaseClas
      p_nReplAmtx = 0
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "cGAwystat") = pxeReplaced Then
            p_nReplAmtx = p_nReplAmtx + .Detail(pnCtr, "nSelPrice")
         End If
      Next
   End With
End Sub

Private Sub Class_Initialize()
   Set p_oSPSales = New clsSPPOSBranch
   pbShowMsg = True
End Sub

Private Sub Class_Terminate()
   Set p_oSPSales = Nothing
End Sub

Private Sub p_oBaseClas_InitDetail(ItemNo As Integer)
   With p_oBaseClas
      If .EditMode = xeModeAddNew Then
         .Detail(ItemNo, "nEntryNox") = ItemNo + 1
      Else
         .Detail(ItemNo, "nEntryNox") = 0
      End If
      .Detail(ItemNo, "sPartsIDx") = ""
      .Detail(ItemNo, "sDescript") = ""
      .Detail(ItemNo, "nQuantity") = 0
      .Detail(ItemNo, "nGivenxxx") = 0
      .Detail(ItemNo, "cGAwyStat") = pxeAdded
      .Detail(ItemNo, "nQtyOnHnd") = 0
      .Detail(ItemNo, "nSelPrice") = 0#
      .Detail(ItemNo, "nResvOrdr") = 0#
      .Detail(ItemNo, "nBackOrdr") = 0#
'160         .Detail(ItemNo, "nReorderx") = 0#
      .Detail(ItemNo, "nLedgerNo") = 0
      .Detail(ItemNo, "sBarrCode") = ""
   End With
End Sub

Private Sub p_oBaseClas_InitMaster()
   With p_oBaseClas
      .Master("sTransNox") = ""
      .Master("dTransact") = p_oAppDrivr.ServerDate()
      .Master("sDRNoxxxx") = ""
      .Master("nTranTotl") = 0
      .Master("nAmtPaidx") = 0
      .Master("sORNoxxxx") = ""
      .Master("sApproved") = ""
      .Master("sRemarksx") = ""
      .Master("cPaymForm") = ""
      .Master("cTranType") = ""
      .Master("cCustType") = ""
      .Master("cTranStat") = xeStateOpen
      .Master("sApplicNo") = ""
      .Master("sClientID") = ""
   
      .Detail(0, "nEntryNox") = 1
      .Detail(0, "sPartsIDx") = ""
      .Detail(0, "sDescript") = ""
      .Detail(0, "nQuantity") = 0
      .Detail(0, "nGivenxxx") = 0
      .Detail(0, "cGAwyStat") = pxeAdded
      .Detail(0, "nQtyOnHnd") = 0
      .Detail(0, "nSelPrice") = 0#
      .Detail(0, "nResvOrdr") = 0#
      .Detail(0, "nBackOrdr") = 0#
'260         .Detail(0, "nReorderx") = 0#
      .Detail(0, "nLedgerNo") = 0
      .Detail(0, "sBarrCode") = ""
   End With
   
   p_sFullName = ""
   p_sAddressx = ""
   
   pbHasSPRep = False
   
   Call p_oSPSales.NewTransaction
End Sub

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      p_sFullName = .Master("xFullName")
      p_sAddressx = .Master("xAddressx")
   
      For pnCtr = 0 To .ItemCount
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "nQtyOnHnd") + .Detail(pnCtr, "nGivenxxx")
      Next
      Call computeReplacement
      
      Call p_oSPSales.NewTransaction
      If p_nReplAmtx > 0 Then
         pbHasSPRep = p_oSPSales.LoadByCondition("sReferNox = " & strParm(.Master("sTransNox")) & _
               " AND sSourceCd = " & strParm(pxeMCSales))
         If pbHasSPRep = False Then
            Call p_oSPSales.NewTransaction
         End If
      End If
   End With
End Sub

'kalyptus - 2022.01:35am
'Add the promo giveaway based on the MC Sales Promo of the period to the regular giveaway
Private Sub loadPromoGiveaway(loGiveAway As Recordset, ByVal lsModelID As String, ByVal lnQty As Integer)
   Dim lsSQL As String
   Dim lsSQx As String
   Dim loRS As Recordset
   Dim loRx As Recordset
   
   'mac 2022-10-22
   '  temporarily for cash, cash balance and installment only
   If Not (p_cPaymForm = "0" Or p_cPaymForm = "1" Or p_cPaymForm = "2") Then Exit Sub
      
   lsSQL = "SELECT" & _
               "  a.sPartsIDx" & _
               ", b.sDescript" & _
               ", a.nQuantity" & _
               ", b.nSelPrice" & _
               ", c.nQtyOnHnd" & _
               ", c.nResvOrdr" & _
               ", c.nBackOrdr" & _
               ", c.nReorderx" & _
               ", c.nLedgerNo" & _
               ", b.sBarrCode" & _
            " FROM MC_Sales_Promo_Giveaway a" & _
               ", MC_Sales_Promo_Model d" & _
               ", MC_Sales_Promo_Master e" & _
               ", Spareparts b" & _
               ", SP_Inventory c"
   lsSQL = lsSQL & _
            " WHERE a.sTransNox = d.sTransNox" & _
              " AND d.sModelIDx = " & strParm(lsModelID) & _
              " AND a.sTransNox = e.sTransNox" & _
              " AND e.cTranStat IN ('1', '2')" & _
              " AND " & dateParm(p_oBaseClas.Master("dTransact")) & " BETWEEN e.dPromoFrm AND e.dPromoTru" & _
              " AND a.sPartsIDx = b.sPartsIDx" & _
              " AND b.sPartsIDx = c.sPartsIDx" & _
              " AND (a.cSlesType = '2' OR a.cSlesType = " & strParm(IIf(p_cPaymForm = "2", "1", "0")) & ")" & _
              " AND c.cRecdStat = " & strParm(xeRecStateActive) & _
              " AND c.sBranchCd = " & strParm(p_sBranchCd)
              
   If LCase(p_oAppDrivr.ProductID) = "integsys" Then
      lsSQx = "SELECT IFNULL(cShopType, '0') cShopType FROM Branch_Others WHERE sBranchCd = " & strParm(p_oAppDrivr.BranchCode)
      Set loRx = New Recordset
      
      loRx.Open lsSQx, p_oAppDrivr.Connection, , , adCmdText
      Set loRx.ActiveConnection = Nothing
      
      If Not loRx.EOF Then
         lsSQL = AddCondition(lsSQL, "IFNULL(e.cShopType, '0') = " & strParm(loRx("cShopType")))
      Else
         lsSQL = AddCondition(lsSQL, "0=1")
      End If
   End If
              
   Debug.Print lsSQL
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
   'Do While loRS.EOF = False
   Do Until loRS.EOF
      loGiveAway.Find "sPartsIDx = " & strParm(loRS("sPartsIDx")), 0, adSearchForward, 1
      
      If loGiveAway.EOF Then
         loGiveAway.AddNew
         loGiveAway("sPartsIDx") = loRS("sPartsIDx")
         loGiveAway("sDescript") = loRS("sDescript")
         loGiveAway("nQuantity") = loRS("nQuantity") * lnQty
         loGiveAway("nQtyOnHnd") = loRS("nQtyOnHnd")
         loGiveAway("nSelPrice") = loRS("nSelPrice")
         loGiveAway("nResvOrdr") = loRS("nResvOrdr")
         loGiveAway("nBackOrdr") = loRS("nBackOrdr")
         loGiveAway("nLedgerNo") = loRS("nLedgerNo")
         loGiveAway("sBarrCode") = loRS("sBarrCode")
      Else
         loGiveAway("nQuantity") = loGiveAway("nQuantity") + _
                                    (loRS("nQuantity") * lnQty)
      End If
      
      loRS.MoveNext
   Loop
End Sub
