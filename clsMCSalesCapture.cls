VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCSalesCapture"
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
'     Motorcycle Sales
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 10/09/2004 11:21 pm ]
'     Start creating this object.
'  XerSys [ 02/21/2005 11:26 am ]
'     Adjust this object to fit to the new implementation of the application driver
'  XerSys [ 11/02/2005 12:56 pm ]
'     Unintegrate the credit application from the mc sales. MC sales will create the
'        transaction credit application, setting it to approved. Then if the application
'        is available, just update it.
'  XerSys [ 11/11/2005 10:28 am ]
'     Adjust the saving of motorcycle to accomodate same model and brand in a single
'        transaction.
'  XerSys [ 01/13/2005 09:53 am ]
'     Add the replaced motorcycle return to the branch...
'  XerSys [ 02/01/2006 03:04 am ]
'     Add the Deed of Sale Expense saving...
'  XerSys [ 04/11/2006 10:38 am ]
'     Add some validation, to lessen user input error...
'     Update the originating branch in LR Ledger...
'  XerSys [ 05/15/2006 10:54 am ]
'     Sold impounded motorcycle must transfer its former account to history. This will
'        remove the old account from the search record of receipt master...
'  XerSys [ 05/23/2006 01:11 pm ]
'     Add sales invoice field...
'  XerSys [ 06/28/2006 09:24 am ]
'     Repo Sales with registration processing fee must have a deed of sale expenses entry
'        regardless of amount paid...
'  XerSys [ 11/09/2006 09:47 am ]
'     Add Term and Due Date field to fully implement MC AR for Institional Customers
'  Kalyptus [ 12/11/2006 03:28 pm]
'     Implements the BugTracker
'  Kalyptus [ 04/02/2007 09:43 am]
'     Update the p_oBase_Clas to conform it with the clsMasterDetail specification
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Type CheckInfo
   CheckNox As String
   AcctNoxx As String
   BankName As String
   CheckDte As Date
   Amountxx As Double
End Type

Private Type ReceiptInfo
   TransNox As String
   ORNoxxxx As String
   TranTotl As Double
   CashAmtx As Double
   CheckAmt As Double
   Remarksx As String
End Type

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1
Private p_oMCARAct As Recordset
Private p_axChecks(4) As CheckInfo
Private p_axReceipt As ReceiptInfo

Private p_sBranchCd As String
Private p_cTranStat As xeTransactionStatus
Private p_sFullName As String
Private p_sAddressx As String
Private p_sActQuery As String
Private p_bGiveAway As Boolean

Private p_sTownCity As String
Private p_sCollectr As String
Private p_sManagerx As String

Private pbModified As Boolean
Private pbInitTran As Boolean
Private pnCtr As Long

Private pbChecksOK As Boolean
Private pbCreatApp As Boolean
Private pbRegisOK As Boolean
Private p_sRegisNox As String

Public Event WillSave(Cancel As Boolean)
Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)
Public Event AccountRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get ARAccount(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xcollectr"
      ARAccount = p_sCollectr
   Case 4, "xmanagerx" & _
      ARAccount = p_sManagerx
   Case Else
      ARAccount = p_oMCARAct(Index)
   End Select
End Property

Property Let ARAccount(ByVal Index As Variant, ByVal Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "sacctnmbr"
      p_oMCARAct(Index) = Value
   Case 1, "sapplicno"
   Case 2, "sroutenme"
      getRoute Value, False
   Case 3, 4, 5, "xcollectr", "xmanagerx", "xsbranchx"
   Case 6, "dfirstpay"
      p_oMCARAct(Index) = Value
      computeDueDate
   Case 7, "nacctterm"
      p_oMCARAct(Index) = Value
      computeDueDate
      computeTerm
   Case 8, 9, 13, 16, 17, 18, "dduedatex", "ngrossprc", "nmonamort", "scolltrid", "smanagrid", "srouteidx"
   Case 10, 11, 12, "ndownpaym", "ncashbalx", "npnvaluex"
      p_oMCARAct(Index) = Value
      computeTerm
   Case 14, 15, "npenaltyx", "nrebatesx"
      p_oMCARAct(Index) = Value
   Case 25, "cloantype"
      p_oMCARAct(Index) = Value
   End Select
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get Checks(ByVal Row As Integer, ByVal Index As Variant) As Variant
   If Row > 4 Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
      Case 0, "schecknox"
         Checks = p_axChecks(Row).CheckNox
      Case 1, "sacctnoxx"
         Checks = p_axChecks(Row).AcctNoxx
      Case 2, "sbankname"
         Checks = p_axChecks(Row).BankName
      Case 3, "dcheckdte"
         Checks = p_axChecks(Row).CheckDte
      Case 4, "namountxx"
         Checks = p_axChecks(Row).Amountxx
   End Select
End Property

Property Let Checks(ByVal Row As Integer, ByVal Index As Variant, ByVal Value As Variant)
   If Row > 4 Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "schecknox"
      p_axChecks(Row).CheckNox = Value
   Case 1, "sacctnoxx"
      p_axChecks(Row).AcctNoxx = Value
   Case 2, "sbankname"
      p_axChecks(Row).BankName = Value
   Case 3, "dcheckdte"
      p_axChecks(Row).CheckDte = Value
   Case 4, "namountxx"
      p_axChecks(Row).Amountxx = Value
   End Select
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
   Case 1, "sengineno"
      getDetail Row, Value, False
   Case 5, "sinstypid"
      getInsurance Row, Value, False
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get GiveAway() As Boolean
   GiveAway = p_bGiveAway
End Property

Property Let GiveAway(Value As Boolean)
   p_bGiveAway = Value
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   With p_oBaseClas
      If Not IsNumeric(Index) Then Index = LCase(Index)
      Select Case Index
      Case 3, "xfullname"
         p_sFullName = Value
      Case 4, "xaddressx"
         p_sAddressx = Value
      Case 6, "namtpaidx"
         .Master(Index) = Value
         
         If .Master("sApplicNo") <> Empty Then
            p_oMCARAct("nDownPaym") = .Master(Index) - .Detail(0, "nInsAmtxx")
            p_oMCARAct("nDownTotl") = .Master(Index) - .Detail(0, "nInsAmtxx")
            If .Master("sApplicNo") = "N-O-N-E" Or pbCreatApp Then
               computeDefault
            Else
               computeTerm
            End If
         End If
      Case 13, "sapplicno"
         If Value = "" Then
            .Master(Index) = Value
         ElseIf Value = "N-O-N-E" Then
            createAccount
         Else
            NewAccount Value
         End If
      Case 21, "stermidxx"
         getTerm Value, False
      Case Else
         .Master(Index) = Value
      End Select
   End With
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get Receipt(ByVal Index As Variant) As Variant
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
      Case 0, "stransnox"
         Receipt = p_axReceipt.TransNox
      Case 1, "sornoxxxx"
         Receipt = p_axReceipt.ORNoxxxx
      Case 2, "ntrantotl"
         Receipt = p_axReceipt.TranTotl
      Case 3, "ncashamtx"
         Receipt = p_axReceipt.CashAmtx
      Case 4, "ncheckamt"
         Receipt = p_axReceipt.CheckAmt
      Case 5, "sremarksx"
         Receipt = p_axReceipt.Remarksx
   End Select
End Property

Property Let Receipt(ByVal Index As Variant, ByVal Value As Variant)
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
      Case 0, "stransnox"
      Case 1, "sornoxxxx"
         p_axReceipt.ORNoxxxx = Value
      Case 2, "ntrantotl"
         p_axReceipt.TranTotl = Value
      Case 3, "ncashamtx"
         p_axReceipt.CashAmtx = Value
      Case 4, "ncheckamt"
         p_axReceipt.CheckAmt = Value
      Case 5, "sremarksx"
         p_axReceipt.Remarksx = Value
   End Select
End Property

Property Let TransStatus(ByVal Value As xeTransactionStatus)
   p_cTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "InitTransaction"
   'On Error Goto errProc
   InitTransaction = False

   If isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   Set p_oBaseClas = New clsMasterDetail
   
   With p_oBaseClas
      Set .AppDriver = p_oAppDrivr
      .MasterTable = "MC_SO_Master"
      .DetailTable = "MC_SO_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDRNoxxxx" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", a.nTranTotl" & _
                        ", a.nAmtPaidx" & _
                        ", a.sORNoxxxx" & _
                        ", a.sApproved" & _
                        ", a.sRemarksx" & _
                        ", a.cPaymForm" & _
                        ", a.cTranType" & _
                        ", a.cCustType" & _
                        ", a.sApplicNo" & _
                        ", a.cTranStat" & _
                        ", a.sClientID" & _
                        ", a.sModified" & _
                        ", a.dModified"
      .MasterQuery = .MasterQuery & _
                        ", b.sCompnyNm" & _
                        ", c.sTownIDxx" & _
                        ", a.sSalesInv" & _
                     " FROM " & .MasterTable & " a" & _
                        ", Client_Master b" & _
                        ", TownCity c" & _
                        ", Province d" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = c.sTownIDxx" & _
                        " AND c.sProvIDxx = d.sProvIDxx" & _
                        IIf(p_cTranStat <> xeStateUnknown, _
                              " AND a.cTranStat = " & p_cTranStat, "")
                        
      .DetailQuery = "SELECT" & _
                        "  a.nEntryNox" & _
                        ", b.sEngineNo" & _
                        ", b.sFrameNox" & _
                        ", c.sModelNme" & _
                        ", d.sColorNme" & _
                        ", e.sInsTypNm" & _
                        ", a.nInsAmtxx" & _
                        ", a.nUnitPrce" & _
                        ", a.sWarrntNo" & _
                        ", a.sRemarksx" & _
                        ", a.sReplMCID" & _
                        ", a.nReplAmtx" & _
                        ", a.nReplIAmt" & _
                        ", f.sMCInvIDx sRMCInvID" & _
                        ", a.cMotorNew" & _
                        ", a.cMotorUnt" & _
                        ", a.sTransNox" & _
                        ", a.dModified" & _
                        ", b.sMCInvIDx" & _
                        ", a.sSerialID" & _
                        ", a.sInsTypID" & _
                        ", e.cRegisTyp"
'      .DetailQuery = .DetailQuery & _
'                     " FROM " & .DetailTable & " a" & _
'                           " LEFT JOIN Insurance_Type e" & _
'                              " ON a.sInsTypID = e.sInsTypID" & _
'                           "  LEFT JOIN MC_Serial b" & _
'                                 " JOIN MC_Model c" & _
'                                    " ON b.sModelIDx = c.sModelIDx" & _
'                                 " JOIN Color d" & _
'                                    " ON b.sColorIDx = d.sColorIDx" & _
'                              " ON a.sSerialID = b.sSerialID" & _
'                           " LEFT JOIN MC_Serial f" & _
'                              " ON a.sReplMCID = f.sSerialID" & _
'                     " ORDER BY a.nEntryNox"
      .DetailQuery = .DetailQuery & _
                     " FROM " & .DetailTable & " a" & _
                           " LEFT JOIN Insurance_Type e ON a.sInsTypID = e.sInsTypID" & _
                           " LEFT JOIN MC_Serial b ON a.sSerialID = b.sSerialID" & _
                           " LEFT JOIN MC_Model c ON b.sModelIDx = c.sModelIDx" & _
                           " LEFT JOIN Color d ON b.sColorIDx = d.sColorIDx" & _
                           " LEFT JOIN MC_Serial f ON a.sReplMCID = f.sSerialID" & _
                     " ORDER BY a.nEntryNox"
   
      .Reference = "a.sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sDRNoxxxx" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND LEFT(a.sTransNox, 2) = " & strParm(p_sBranchCd) & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sDRNoxxxx»xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sDRNoxxxx»CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "DR No»Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@»@»@»@@@@-@@@@@@»MMMM DD, YYYY"
      
      .BrowseQuery(15) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) AS xFullName" & _
                           ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(15) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(15) = "Client ID»Customer Name»Address"
      .BrowseCriteria(15) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"
      
'      ' Query statement for the term field
'      .BrowseQuery(23) = "SELECT" & _
'                        "  sTermIDxx" & _
'                        ", sTermName" & _
'                        ", nTermDays" & _
'                        ", nDiscDays" & _
'                        ", nDiscount" & _
'                     " FROM Term" & _
'                     " WHERE cRecdStat = " & xeRecStateActive
'      .BrowseColumn(23) = "sTermIDxx»sTermName»nTermDays»nDiscDays»nDiscount"
'      .BrowseTitle(23) = "Code»Term Name»Term»Disc Days»Disc"
'      .BrowseFormat(23) = "@»@»0 Day/s»0 Day/s»0.00"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sSerialID" & _
                                 ", a.sEngineNo" & _
                                 ", a.sFrameNox" & _
                                 ", b.sModelNme" & _
                                 ", c.sColorNme" & _
                                 ", d.nSelPrice" & _
                                 ", a.sMCInvIDx" & _
                                 ", a.cSoldStat" & _
                              " FROM MC_Serial a" & _
                                 ", MC_Model b" & _
                                 ", Color c" & _
                                 ", MC_Inventory d" & _
                              " WHERE a.sModelIDx = b.sModelIDx" & _
                                 " AND a.sColorIDx = c.sColorIDx" & _
                                 " AND a.sMCInvIDx = d.sMCInvIdx" & _
                                 " AND a.sBranchCd = d.sBranchCd" & _
                                 " AND a.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND a.cLocation = " & strParm(xeLocBranch)
      .BrowseDetailColumn(0) = "sEngineNo»sFrameNox»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Engine No»Frame No»Model»Color"
      .BrowseDetailCriteria(0) = "a.sEngineNo»a.sFrameNox»b.sModelNme»c.sColorNme"
   End With
   
   p_sActQuery = "SELECT" & _
                     "  h.sAcctNmbr" & _
                     ", h.sApplicNo" & _
                     ", c.sRouteNme" & _
                     ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) AS xCollectr" & _
                     ", CONCAT(e.sLastName, ', ', e.sFrstName, ' ', e.sMiddName) AS xManagerx" & _
                     ", f.sBranchNm xSBranchx" & _
                     ", h.dFirstPay" & _
                     ", a.nAcctTerm" & _
                     ", h.dDueDatex" & _
                     ", a.nGrossPrc" & _
                     ", a.nDownPaym" & _
                     ", h.nCashBalx" & _
                     ", a.nPNValuex" & _
                     ", a.nMonAmort" & _
                     ", h.nPenaltyx" & _
                     ", h.nRebatesx" & _
                     ", c.sCollctID" & _
                     ", c.sManagrID" & _
                     ", g.sRouteIDx" & _
                     ", h.nDownTotl" & _
                     ", h.nABalance"

   p_sActQuery = p_sActQuery & _
                     ", a.cTranStat" & _
                     ", f.sBranchCd" & _
                     ", b.sTownIDxx" & _
                     ", a.sRemarksx" & _
                     ", h.cLoanType" & _
                  " FROM MC_Credit_Application a" & _
                        " LEFT JOIN MC_AR_Master h" & _
                           " ON a.sTransNox = h.sApplicNo" & _
                     ", Client_Master b" & _
                     ", Route_Area c" & _
                     ", Employee_Master d" & _
                     ", Employee_Master e" & _
                     ", Branch f" & _
                     ", Route_Area_Town g" & _
                     ", TownCity i" & _
                     ", Province j" & _
                  " WHERE a.sClientID = b.sClientID" & _
                     " AND b.sTownIDxx = g.sTownIDxx" & _
                     " AND g.sRouteIDx = c.sRouteIDx" & _
                     " AND c.sCollctID = d.sEmployID" & _
                     " AND c.sManagrID = e.sEmployID" & _
                     " AND c.sBranchCd = f.sBranchCd" & _
                     " AND b.sTownIDxx = i.sTownIDxx" & _
                     " AND i.sProvIDxx = j.sProvIDxx"

   pbInitTran = InitTransaction
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "NewTransaction"
   'On Error Goto errProc

   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction
endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SaveTransaction"
   'On Error Goto errProc

   If pbInitTran Then SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "OpenTransaction"
   'On Error Goto errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "DeleteTransaction"
   'On Error Goto errProc

   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:

   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsRecord As String
   Dim Value As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsOldProc = "SearchTransaction"
   'On Error Goto errProc
   SearchTransaction = False
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      Value = sSearch
      lsField = "a.sDRNoxxxx"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
            lsDescript = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
         End If
      End If
   Else
      'Perform this if there is no entry for DR Number and Customer Name
      'As requested by Jen Lancen de Magiba
      Value = "%"
      lsField = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
      lsDescript = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
   End If
   
   SearchTransaction = p_oBaseClas.SearchTransaction(Value, lsField, lsDescript)
   
endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "PostTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "CancelTransaction"
   'On Error Goto errProc
   
   If pbInitTran Then
      With p_oBaseClas
         If pbRegisOK = False Then
            MsgBox "Registration of the motorcycle included in the sales is on process!!!" & _
                     vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
            Exit Function
         ElseIf pbChecksOK = False Then
            MsgBox "Check payment of this transaction has been deposited!!!" & _
                     vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
            Exit Function
         End If
         CancelTransaction = .CancelTransaction()
      End With
   End If

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "CloseTransaction"
   'On Error Goto errProc
      
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)




endProc:

   Exit Function
errProc:
    ShowError lsOldProc & "( " & sTransNo & " )"
End Function

Function AddDetail() As Boolean
   AddDetail = p_oBaseClas.AddDetail
   
   pnCtr = p_oBaseClas.ItemCount - 1
   With p_oBaseClas
      .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      .Detail(pnCtr, "sEngineNo") = ""
      .Detail(pnCtr, "sFrameNox") = ""
      .Detail(pnCtr, "sModelNme") = ""
      .Detail(pnCtr, "sColorNme") = ""
      .Detail(pnCtr, "sInsTypNm") = ""
      .Detail(pnCtr, "nInsAmtxx") = 0#
      .Detail(pnCtr, "nUnitPrce") = 0#
      .Detail(pnCtr, "sWarrntNo") = ""
      .Detail(pnCtr, "sRemarksx") = ""
      .Detail(pnCtr, "cMotorNew") = ""
      .Detail(pnCtr, "sMCInvIDx") = ""
      .Detail(pnCtr, "sSerialID") = ""
      .Detail(pnCtr, "sInsTypID") = ""
      .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
      .Detail(pnCtr, "cMotorUnt") = 0
      .Detail(pnCtr, "sReplMCID") = ""
      .Detail(pnCtr, "nReplAmtx") = 0
      .Detail(pnCtr, "nReplIAmt") = 0
      .Detail(pnCtr, "sRMCInvID") = 0
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SearchDetail"
   'On Error Goto errProc
   SearchDetail = False
   
   If pbInitTran = False Then GoTo endProc
   
   Index = LCase(Index)
   Select Case Index
   Case 1, "sengineno"
      SearchDetail = getDetail(Row, Value, True)
   Case 5, "sinstypnm"
      SearchDetail = getInsurance(Row, Value, True)
   End Select
   
endProc:

   Exit Function
errProc:



    ShowError lsOldProc & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
   GoTo endProc
End Function

Function SearchMaster(ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean



   Dim lsOldProc As String
   
   lsOldProc = "SearchMaster"
   'On Error Goto errProc
      
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      SearchMaster = GetCustomer(Value, True)
   Case 13, "sapplicno"
      SearchMaster = getApplication
   Case 21, "stermidxx"
      SearchMaster = getTerm(Value, True)
   End Select




endProc:

   Exit Function
errProc:
    ShowError lsOldProc & "( " & Index _
                        & ", " & Value & " )"
End Function

Function SearchAccount(ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "SearchAccount"
   'On Error Goto errProc
      
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 7, "sroutenme"
      SearchAccount = getRoute(Value, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & Index _
                        & ", " & Value & " )"
End Function

Function UpdateTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "UpdateTransaction"
   'On Error Goto errProc
   
   With p_oBaseClas
      If OpenTransaction(.Master("sTransNox")) = False Then Exit Function
      
      If pbRegisOK = False Then
         MsgBox "Registration of the motorcycle included in the sales is on process!!!" & _
                  vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
         Exit Function
      ElseIf pbChecksOK = False Then
         MsgBox "Check payment of this transaction has been deposited!!!" & _
                  vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
         Exit Function
      End If
      
      .EditMode = xeModeUpdate
   End With
   
   UpdateTransaction = True

endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function saveAccount() As Boolean
   Dim lsSQL As String
   Dim lsDetailSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   Dim lnCashTotl As Double
   Dim lnDownTotl As Double
   Dim lnEditMode As xeEditMode

   lsOldProc = "saveAccount"
   'On Error Goto errProc
   saveAccount = False

   With p_oBaseClas
      If .Master("sApplicNo") = Empty And .Master("nTranTotl") > .Master("nAmtPaidx") And .Master("cPaymForm") <> 3 Then
         MsgBox "Invalid Credit Application Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lnEditMode = xeModeReady
      If .EditMode = xeModeAddNew Or .EditMode = xeModeUpdate Then
         If .Master("sApplicNo") <> Empty Then lnEditMode = xeModeAddNew
      End If
   End With
   
   If lnEditMode = xeModeReady Then
      saveAccount = True
      GoTo endProc
   End If
   
   With p_oBaseClas
      lnDownTotl = .Master("nAmtPaidx") - .Detail(0, "nInsAmtxx")
      lnCashTotl = 0#
   End With
   
   With p_oMCARAct
      ' before saving account, check if account contains valid information
      If .Fields("nMonAmort") = 0# And .Fields("nCashBalx") = 0# Then
         MsgBox "Invalid Account Term Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If (.Fields("nAcctTerm") = 0 And .Fields("nCashBalx") = 0) _
         Or (.Fields("nPNValuex") > 0 And .Fields("nCashBalx") > 0) Then
         MsgBox "Invalid Account Term Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If DateDiff("d", p_oBaseClas.Master("dTransact"), .Fields("dFirstPay")) > 60 Then
         MsgBox "Invalid First Payment Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If .Fields("sRouteIDx") = Empty Then
         MsgBox "Invalid Collection Route Detected!!!" & vbCrLf & _
                  "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If p_oBaseClas.Master("cTranType") <> 2 Then
         lsSQL = "INSERT INTO MC_AR_Master " & _
                     "( sAcctNmbr" & _
                     ", sBranchCd" & _
                     ", sClientID" & _
                     ", sApplicNo" & _
                     ", sRouteIDx" & _
                     ", sRemarksx" & _
                     ", sExAcctNo" & _
                     ", sSerialID" & _
                     ", dPurchase" & _
                     ", nGrossPrc" & _
                     ", nPNValuex" & _
                     ", nDownPaym" & _
                     ", nCashBalx" & _
                     ", dFirstPay" & _
                     ", nAcctTerm" & _
                     ", dDueDatex" & _
                     ", nMonAmort" & _
                     ", nPenaltyx" & _
                     ", nRebatesx" & _
                     ", nLastPaym"
   
         lsSQL = lsSQL & _
                     ", dLastPaym" & _
                     ", nPaymTotl" & _
                     ", nPenTotlx" & _
                     ", nRebTotlx" & _
                     ", nDebtTotl" & _
                     ", nCredTotl" & _
                     ", nAmtDuexx" & _
                     ", nABalance" & _
                     ", nDownTotl" & _
                     ", nCashTotl" & _
                     ", nDelayAvg" & _
                     ", cRatingxx" & _
                     ", cAcctstat" & _
                     ", cMotorNew" & _
                     ", cLoanType" & _
                     ", cActivexx" & _
                     ", nLgrEntry" & _
                     ", sModified" & _
                     ", dModified"
         lsSQL = lsSQL & " ) VALUES ( " & _
                     strParm(.Fields("sAcctNmbr")) & _
                     ", " & strParm(Format(.Fields("sBranchCd"), "00")) & _
                     ", " & strParm(p_oBaseClas.Master("sClientID")) & _
                     ", " & strParm(.Fields("sApplicNo")) & _
                     ", " & strParm(.Fields("sRouteIDx")) & _
                     ", " & strParm("") & _
                     ", " & strParm("") & _
                     ", " & strParm(p_oBaseClas.Detail(0, "sSerialID")) & _
                     ", " & dateParm(p_oBaseClas.Master("dTransact")) & _
                     ", " & .Fields("nGrossPrc") & _
                     ", " & .Fields("nPNValuex") & _
                     ", " & .Fields("nDownPaym") & _
                     ", " & .Fields("nCashBalx") & _
                     ", " & dateParm(.Fields("dFirstPay")) & _
                     ", " & .Fields("nAcctTerm") & _
                     ", " & dateParm(.Fields("dDueDatex")) & _
                     ", " & .Fields("nMonAmort") & _
                     ", " & .Fields("nPenaltyx") & _
                     ", " & .Fields("nRebatesx") & _
                     ", 0.00"
         lsSQL = lsSQL & _
                     ", " & dateParm(p_oBaseClas.Master("dTransact")) & _
                     ", 0.00" & _
                     ", 0.00" & _
                     ", 0.00" & _
                     ", 0.00" & _
                     ", 0.00" & _
                     ", 0.00" & _
                     ", " & .Fields("nABalance") & _
                     ", " & lnDownTotl & _
                     ", 0.00" & _
                     ", 0.00" & _
                     ", " & strParm("") & _
                     ", " & strParm(xeActStatActive) & _
                     ", " & strParm(p_oBaseClas.Detail(0, "cMotorNew")) & _
                     ", " & strParm(.Fields("cLoanType")) & _
                     ", " & strParm(xeYes) & _
                     ", " & 1 & _
                     ", " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                     ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
                     
         lsDetailSQL = "INSERT INTO MC_AR_Ledger " & _
                           "( sAcctNmbr" & _
                           ", sBranchCd" & _
                           ", nEntryNox" & _
                           ", dTransact" & _
                           ", cOffPaymx" & _
                           ", sCollIDxx" & _
                           ", sORNoxxxx" & _
                           ", cTrantype" & _
                           ", sRemarksx" & _
                           ", nTranAmtx" & _
                           ", nDebitAmt" & _
                           ", nOthersxx" & _
                           ", nRebatesx" & _
                           ", nABalance" & _
                           ", nMonDelay" & _
                           ", dModified"
         lsDetailSQL = lsDetailSQL & " ) VALUES (" & _
                           strParm(.Fields("sAcctNmbr")) & _
                           ", " & strParm(Left(.Fields("sAcctNmbr"), 2)) & _
                           ", " & 1 & _
                           ", " & dateParm(p_oBaseClas.Master("dTransact")) & _
                           ", " & strParm(1) & _
                           ", " & strParm("") & _
                           ", " & strParm(p_oBaseClas.Master("sORNoxxxx")) & _
                           ", " & strParm("d") & _
                           ", " & strParm("Initial DownPayment") & _
                           ", " & lnDownTotl & _
                           ", " & 0# & _
                           ", " & 0# & _
                           ", " & 0# & _
                           ", " & .Fields("nPNValuex") + .Fields("nDownPaym") - lnDownTotl & _
                           ", " & 0# & _
                           ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      Else
         lsSQL = "UPDATE MC_AR_Master SET" & _
                     "  sSerialID = " & strParm(p_oBaseClas.Detail(0, "sSerialID")) & _
                     ", nGrossPrc = " & .Fields("nGrossPrc") & _
                     ", nPNValuex = " & .Fields("nPNValuex") & _
                     ", nDownPaym = " & .Fields("nDownPaym") & _
                     ", nCashBalx = " & .Fields("nCashBalx") & _
                     ", dFirstPay = " & dateParm(.Fields("dFirstPay")) & _
                     ", nAcctTerm = " & .Fields("nAcctTerm") & _
                     ", dDueDatex = " & dateParm(.Fields("dDueDatex")) & _
                     ", nMonAmort = " & .Fields("nMonAmort") & _
                     ", nRebatesx = " & .Fields("nRebatesx") & _
                     ", nABalance = " & .Fields("nABalance") & _
                     ", nDownTotl = nDownTotl + " & p_oBaseClas.Master("nAmtPaidx") - p_oBaseClas.Detail(0, "nInsAmtxx") & _
                     ", sModified = " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                  " WHERE sAcctNmbr = " & strParm(.Fields("sAcctNmbr"))
         
         lsDetailSQL = "UPDATE MC_AR_Ledger SET" & _
                           "  nTranAmtx = nTranAmtx + " & p_oBaseClas.Master("nAmtPaidx") - p_oBaseClas.Detail(0, "nInsAmtxx") & _
                           ", nABalance = " & .Fields("nPNValuex") + .Fields("nDownPaym") - lnDownTotl & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                        " WHERE sAcctNmbr = " & strParm(.Fields("sAcctNmbr")) & _
                           " AND nEntryNox = 1"
      End If
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
   
      If Row <= 0 Then
         MsgBox "Unable to Save MC AR Account Info!!!" & lsSQL, vbCritical, "Warning"
         GoTo endProc
      End If
   
      p_oAppDrivr.Connection.Execute lsDetailSQL, Row, adCmdText
   
      If Row <= 0 Then
         MsgBox "Unable to Save MC AR Detail Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If .Fields("sApplicNo") <> "N-O-N-E" Then
         lsSQL = "UPDATE MC_Credit_Application SET" & _
                     " cTranStat = " & strParm(xeStateUnknown) & _
                  " WHERE sTransNox = " & strParm(.Fields("sApplicNo"))
         
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
         If Row <= 0 Then
            MsgBox "Unable to Update MC Credit Application Info!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With
   
   saveAccount = True
   
endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, "clsMCSalesCapture", lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function GetAccountQuery() As String
   GetAccountQuery = "SELECT" & _
                     "  h.sAcctNmbr" & _
                     ", h.sApplicNo" & _
                     ", c.sRouteNme" & _
                     ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) AS xCollectr" & _
                     ", CONCAT(e.sLastName, ', ', e.sFrstName, ' ', e.sMiddName) AS xManagerx" & _
                     ", f.sBranchNm xSBranchx" & _
                     ", h.dFirstPay" & _
                     ", a.nAcctTerm" & _
                     ", h.dDueDatex" & _
                     ", a.nGrossPrc" & _
                     ", a.nDownPaym" & _
                     ", h.nCashBalx" & _
                     ", a.nPNValuex" & _
                     ", a.nMonAmort" & _
                     ", h.nPenaltyx" & _
                     ", h.nRebatesx" & _
                     ", c.sCollctID" & _
                     ", c.sManagrID" & _
                     ", g.sRouteIDx"
   
   GetAccountQuery = GetAccountQuery & _
                     ", a.cTranStat" & _
                     ", f.sBranchCd" & _
                     ", b.sTownIDxx" & _
                     ", a.sRemarksx" & _
                     ", h.cLoanType" & _
                     ", h.nDownTotl" & _
                     ", h.nABalance"

   GetAccountQuery = GetAccountQuery & _
                  " FROM MC_Credit_Application a" & _
                        " RIGHT JOIN MC_AR_Master h" & _
                           " ON a.sTransNox = h.sApplicNo" & _
                     ", Client_Master b" & _
                     ", Route_Area c" & _
                     ", Employee_Master d" & _
                     ", Employee_Master e" & _
                     ", Branch f" & _
                     ", Route_Area_Town g" & _
                     ", TownCity i" & _
                     ", Province j" & _
                  " WHERE a.sClientID = b.sClientID" & _
                     " AND b.sTownIDxx = g.sTownIDxx" & _
                     " AND g.sRouteIDx = c.sRouteIDx" & _
                     " AND c.sCollctID = d.sEmployID" & _
                     " AND c.sManagrID = e.sEmployID" & _
                     " AND c.sBranchCd = f.sBranchCd" & _
                     " AND b.sTownIDxx = i.sTownIDxx" & _
                     " AND i.sProvIDxx = j.sProvIDxx"
End Function

Function GetCustomer(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsOldProc As String

   lsOldProc = "GetCustomer"
   'On Error Goto errProc
   GetCustomer = False
   
   With p_oBaseClas
      If Value <> "" Then
         If Value = p_sFullName Then GoTo endProc
         
         lsMaster = "CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)"
         If Search Then
            lsMaster = lsMaster & " = " & strParm(Trim(Value))
         Else
            lsMaster = lsMaster & " LIKE " & strParm(Trim(Value) & "%")
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If
      
      lsMaster = .getMaster(15, lsMaster)
      If lsMaster = Empty Then
         If Search = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      .Master("sClientID") = lasMaster(0)
      p_sFullName = lasMaster(1)
      p_sAddressx = lasMaster(2)
   End With
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   
   GetCustomer = True
   
endProc:
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
   p_sFullName = ""
   p_sAddressx = ""
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & Value _
                        & ", " & Search & " )"
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Function getApplication() As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lnReply As Integer
   
   lsOldProc = "GetApplication"
   'On Error Goto errProc
   getApplication = False
   
   With p_oBaseClas
      If .Master("sClientID") = Empty Then GoTo endWithClear
      
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                  ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
               " FROM MC_Credit_Application a" & _
                  ", Client_Master b" & _
                  ", TownCity c" & _
                  ", Province d" & _
               " WHERE a.sClientID = b.sClientID" & _
                  " AND b.sTownIDxx = c.sTownIDxx" & _
                  " AND c.sProvIDxx = d.sProvIDxx" & _
                  " AND b.sClientID = " & strParm(.Master("sClientID"))
      
      If .Master("cTranType") = 2 Then
         ' replacement sales
         ' get the former application
         lsSQL = lsSQL & _
                  " ORDER BY a.dAppliedx DESC"
      Else
         lsSQL = lsSQL & _
                     " AND a.cTranStat <> " & strParm(xeStateUnknown)
      End If
      
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

      If lors.EOF Then
         lnReply = MsgBox("Customer Has no Credit Application!!!" & vbCrLf & _
                           " Click [YES] to create Application," & vbCrLf & _
                           " [NO] to continue anyway," & vbCrLf & _
                           " [CANCEL] to abort entry", vbCritical + vbYesNoCancel, "Confirm")
                           
         If lnReply = vbYes Or lnReply = vbNo Then
            pbCreatApp = lnReply = vbYes
            If createAccount = False Then GoTo endProc
            
            .Master("sApplicNo") = p_oMCARAct("sApplicNo")
            RaiseEvent MasterRetrieved(13)
            getApplication = True
            GoTo endProc
         End If
      ElseIf lors.RecordCount = 1 Then
         If NewAccount(lors("sTransNox")) = False Then GoTo endWithClear
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, lors, "sTransNox»xFullName»xAddressx", _
                              "Application No»Customer»Address")
         
         If lsSQL = Empty Then GoTo endWithClear
         lasMaster = Split(lsSQL, "»")
         If NewAccount(lasMaster(0)) = False Then GoTo endWithClear
      End If
   End With
   
   RaiseEvent MasterRetrieved(13)
   
   getApplication = True
   
endProc:
   Set lors = Nothing

   Exit Function
endWithClear:
   p_oBaseClas.Master("sApplicNo") = ""
   RaiseEvent MasterRetrieved(13)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function getRoute(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lsFilter As String
   Dim lsOldProc As String

   lsOldProc = "getRoute"
   'On Error Goto errProc
   getRoute = False

   With p_oMCARAct
      If Value = .Fields("sRouteNme") Then
         getRoute = True
         GoTo endProc
      End If
   
      lsSQL = "SELECT" & _
                  "  a.sRouteIDx" & _
                  ", a.sRouteNme" & _
                  ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xCollectr" & _
                  ", b.sEmployID" & _
                  ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) AS xManagerx" & _
                  ", d.sEmployID xManagrID" & _
                  ", e.sBranchNm" & _
                  ", e.sBranchCd" & _
               " FROM Route_Area a" & _
                  ", Employee_Master b" & _
                  ", Route_Area_Town c" & _
                  ", Employee_Master d" & _
                  ", Branch e" & _
               " WHERE a.sCollctID = b.sEmployID" & _
                  " AND a.sRouteIDx = c.sRouteIDx" & _
                  " AND a.sManagrID = d.sEmployID" & _
                  " AND a.sBranchCd = e.sBranchCd" & _
                  " AND c.sTownIDxx = " & strParm(.Fields("sTownIDxx"))
   
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then
         MsgBox "Town is not Included in any Area!!!" & vbCrLf & _
                  "Creation of Account is not Allowed!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If lors.RecordCount = 1 Then
         .Fields("sRouteIDx") = lors("sRouteIDx")
         .Fields("sRouteNme") = lors("sRouteNme")
         .Fields("sCollctID") = lors("sEmployID")
         .Fields("sManagrID") = lors("xManagrID")
         .Fields("sBranchCd") = lors("sBranchCd")
         p_sCollectr = lors("xCollectr")
         p_sManagerx = lors("xManagerx")
         GoTo endWithEvent
      End If
         
         lsSQL = KwikBrowse(p_oAppDrivr, lors, "sRouteIDx»sRouteNme»xCollectr", "Route Code»Route Name»Collector", , "a.sRouteIDx»a.sRouteNme»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)")
         If lsSQL <> Empty Then
            lasMaster = Split(lsSQL, "»")
            .Fields("sRouteIDx") = lasMaster(0)
            .Fields("sRouteNme") = lasMaster(1)
            .Fields("sCollctID") = lasMaster(3)
            .Fields("sManagrID") = lasMaster(5)
            .Fields("sBranchCd") = lasMaster(7)
            p_sCollectr = lasMaster(2)
            p_sManagerx = lasMaster(4)
            GoTo endWithEvent
         Else
            lors.MoveFirst
         End If
      .Fields("sRouteIDx") = lors("sRouteIDx")
      .Fields("sCollctID") = lors("sEmployID")
      .Fields("sManagrID") = lors("xManagrID")
      .Fields("sBranchCd") = lors("sBranchCd")
      p_sCollectr = lors("xCollectr")
      p_sManagerx = lors("xManagerx")
   End With
   
endWithEvent:
   RaiseEvent AccountRetrieved(2)
   RaiseEvent AccountRetrieved(3)
   RaiseEvent AccountRetrieved(4)

   getRoute = True
   GoTo endProc
endProc:
   Set lors = Nothing
   Exit Function
errProc:
    ShowError lsOldProc & "( " & Value _
                        & ", " & Search & " )"
   GoTo endProc
End Function

Private Function getTerm(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsOldProc As String

   lsOldProc = "GetTerm"
   'On Error Goto errProc
   
   With p_oBaseClas
      If Value <> "" Then
         If Value = .Master("sTermName") Then
            getTerm = True
            GoTo endProc
         End If
         
         If Search Then
            lsMaster = "sTermName LIKE " & strParm(Trim(Value) & "%") '(previous code of this line) lsMaster = "a.sTermName LIKE " & strParm(Trim(Value) & "%")
         Else
            lsMaster = "sTermName = " & strParm(Trim(Value)) '(previous code of this line) lsMaster = "a.sTermName = " & strParm(Trim(Value))
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If
   
      lsMaster = .getMaster(23, lsMaster)
      If lsMaster = Empty Then
         If Search = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasMaster = Split(lsMaster, "»")
      
      .Master("sTermIDxx") = lasMaster(0)
      .Master("sTermName") = lasMaster(1)
      .Master("dDueDatex") = DateAdd("d", lasMaster(2), .Master("dTransact"))
   End With

   getTerm = True
   
endProc:
   RaiseEvent MasterRetrieved(21)

   Exit Function
endWithClear:
   With p_oBaseClas
      .Master("sTermIDxx") = ""
      .Master("sTermName") = ""
      .Master("dDueDatex") = .Master("dTransact")
   End With
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & Value _
                        & ", " & Search & " )"
End Function

Private Function createAccount() As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lsSellCode As String
   Dim lsCollCode As String
   Dim Row As Integer
   
   lsOldProc = "CreateAccount"
   'On Error Goto errProc
   createAccount = False
   
   lsSQL = "SELECT" & _
                     "  h.sAcctNmbr" & _
                     ", h.sApplicNo" & _
                     ", c.sRouteNme" & _
                     ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) AS xCollectr" & _
                     ", CONCAT(e.sLastName, ', ', e.sFrstName, ' ', e.sMiddName) AS xManagerx" & _
                     ", f.sBranchNm xSBranchx" & _
                     ", h.dFirstPay" & _
                     ", h.nAcctTerm" & _
                     ", h.dDueDatex" & _
                     ", h.nGrossPrc" & _
                     ", h.nDownPaym" & _
                     ", h.nCashBalx" & _
                     ", h.nPNValuex" & _
                     ", h.nMonAmort" & _
                     ", h.nPenaltyx" & _
                     ", h.nRebatesx" & _
                     ", c.sCollctID" & _
                     ", c.sManagrID" & _
                     ", g.sRouteIDx"
   lsSQL = lsSQL & _
                     ", h.nDownTotl" & _
                     ", h.nABalance" & _
                     ", h.cLoanType" & _
                     ", f.sBranchCd" & _
                     ", a.sTownIDxx" & _
                  " FROM TownCity a" & _
                        " LEFT JOIN MC_AR_Master h" & _
                           " ON h.sAcctNmbr = " & strParm("") & _
                     ", Province b" & _
                     ", Route_Area c" & _
                     ", Employee_Master d" & _
                     ", Employee_Master e" & _
                     ", Branch f" & _
                     ", Route_Area_Town g" & _
                  " WHERE a.sProvIDxx = b.sProvIDxx" & _
                     " AND a.sTownIDxx = g.sTownIDxx" & _
                     " AND g.sRouteIDx = c.sRouteIDx" & _
                     " AND c.sCollctID = d.sEmployID" & _
                     " AND c.sManagrID = e.sEmployID" & _
                     " AND c.sBranchCd = f.sBranchCd" & _
                     " AND a.sTownIDxx = " & strParm(p_oBaseClas.Master("sTownIDxx"))
                     
   Set p_oMCARAct = New Recordset
   p_oMCARAct.Open lsSQL, p_oAppDrivr.Connection, , adLockOptimistic, adCmdText
   If p_oMCARAct.EOF Then GoTo endProc
   
   With p_oMCARAct
      If IsNull(p_oMCARAct("sAcctNmbr")) Then
         ' after getting the account no then assign the no to the account
         .Fields("sAcctNmbr") = GetNextCode("MC_AR_Master", "sAcctNmbr", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
         .Fields("dFirstPay") = DateAdd("m", 1, p_oBaseClas.Master("dTransact"))
         .Fields("nGrossPrc") = 0#
         .Fields("nDownPaym") = 0#
         .Fields("nDownTotl") = 0#
         .Fields("nCashBalx") = 0#
         .Fields("nPNValuex") = 0#
         .Fields("nMonAmort") = 0#
         .Fields("nPenaltyx") = 0#
         .Fields("nRebatesx") = 0#
         .Fields("cLoanType") = 0   ' Motorcycle
      End If
      p_sCollectr = .Fields("xCollectr")
      p_sManagerx = .Fields("xManagerx")
      
      If pbCreatApp Then
         .Fields("sApplicNo") = GetNextCode("MC_Credit_Application", "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
         .Fields("nAcctTerm") = 12
         .Fields("dDueDatex") = DateAdd("m", .Fields("nAcctTerm"), p_oBaseClas.Master("dTransact"))
      Else
         .Fields("sApplicNo") = "N-O-N-E"
         .Fields("nAcctTerm") = 0
         .Fields("dDueDatex") = .Fields("dFirstPay")
      End If
      
   End With
   
   createAccount = True
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

' pass the application no of the credit application in order
'   to check the application status.
Private Function NewAccount(ByVal lsApplicNo As String) As Boolean
   Dim lors As New Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim Row As Long
   
   lsOldProc = "NewAccount"
   'On Error Goto errProc
   NewAccount = False
   
   If Not pbInitTran Then GoTo endProc
   If lsApplicNo = Empty Then GoTo endProc
   
   lsSQL = p_sActQuery & _
               " AND a.sTransNox = " & strParm(lsApplicNo)
               
   Set p_oMCARAct = New Recordset
   p_oMCARAct.Open lsSQL, p_oAppDrivr.Connection, , adLockOptimistic, adCmdText
   
   If p_oMCARAct.EOF Then
      MsgBox "Application No Does not Exist!!!" & vbCrLf & _
               "Verify Your Entry then Try Again!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   Select Case p_oMCARAct("cTranStat")
   Case xeStateOpen, xeStateClosed
      MsgBox "Application has Not Yet been Approved!!!" & vbCrLf & _
               "Account Creation is Not Allowed!!!", vbCritical, "Warning"
      GoTo endProc
   Case xeStateCancelled
      MsgBox "Application was Disapproved!!!" & vbCrLf & _
               "Account Creation is Not Allowed!!!", vbCritical, "Warning"
      GoTo endProc
   Case xeStateUnknown
      If p_oBaseClas.Master("cTranType") <> 2 Then
         MsgBox "Account was Already Created from this Application!!!" & vbCrLf & _
                  "Account Re-Creation is Not Allowed!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End Select
   
   ' after getting the account no then assign the no to the account
   If IsNull(p_oMCARAct("sAcctNmbr")) Then
      p_oMCARAct("sAcctNmbr") = GetNextCode("MC_AR_Master", "sAcctNmbr", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
   End If
   p_oMCARAct("nDownTotl") = p_oMCARAct("nDownPaym")
   p_oMCARAct("nCashBalx") = 0#
   p_oMCARAct("nPenaltyx") = 0#
   p_oMCARAct("nRebatesx") = 0#
   p_oMCARAct("dFirstPay") = DateAdd("m", 1, p_oBaseClas.Master("dTransact"))
   p_oMCARAct("sApplicNo") = lsApplicNo
   p_oMCARAct("nABalance") = p_oMCARAct("nGrossPrc") - p_oMCARAct("nDownTotl")
   p_oMCARAct("cLoanType") = 0   ' Motorcycle
   p_oBaseClas.Master("sApplicNo") = lsApplicNo
   
   p_sCollectr = p_oMCARAct("xCollectr")
   p_sManagerx = p_oMCARAct("xManagerx")
   computeDueDate
   NewAccount = True
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsApplicNo & " )"
   GoTo endProc
End Function

Private Function getDetail(ByVal Row As Integer, _
                        ByVal Value As String, _
                        ByVal Search As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String

   lsOldProc = "getDetail"
   'On Error Goto errProc
   getDetail = False
   
   With p_oBaseClas
      If Value <> "" Then
         If Value = .Detail(Row, "sEngineNo") Then GoTo endProc
         If Search Then
            lsDetail = "a.sEngineNo LIKE " & strParm("%" & Trim(Value))
         Else
            lsDetail = "a.sEngineNo = " & strParm(Trim(Value))
         End If
      ElseIf Search = False Then
         GoTo endWithClear
      End If
      
      lsDetail = .getDetail(Row, lsDetail)
      If lsDetail = Empty Then
         If Search = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasDetail = Split(lsDetail, "»")
      .Detail(Row, "sEngineNo") = lasDetail(1)
      .Detail(Row, "sFrameNox") = lasDetail(2)
      .Detail(Row, "sModelNme") = lasDetail(3)
      .Detail(Row, "sColorNme") = lasDetail(4)
      .Detail(Row, "sSerialID") = lasDetail(0)
      .Detail(Row, "nUnitPrce") = lasDetail(5)
      .Detail(Row, "sMCInvIDx") = lasDetail(6)
      .Detail(Row, "cMotorUnt") = 1
      .Detail(Row, "sReplMCID") = ""
      .Detail(Row, "nReplAmtx") = 0#
      .Detail(Row, "nReplIAmt") = 0#
      If lasDetail(7) = Empty Then
         .Detail(Row, "cMotorNew") = xeYes
      Else
         .Detail(Row, "cMotorNew") = IIf(lasDetail(7) = xeNo, 1, 0)
      End If
      
      If .Master("cTranType") = 2 Then
         ' Replacement transaction must return the changed unit
         '  so get the mc info then process the returned unit
         If getReplacedUnit(Row) = False Then GoTo endWithClear
      End If
   End With
   
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(7)
   
   getDetail = True
   
endProc:
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(Row, "sEngineNo") = ""
      .Detail(Row, "sFrameNox") = ""
      .Detail(Row, "sModelNme") = ""
      .Detail(Row, "sColorNme") = ""
      .Detail(Row, "sSerialID") = ""
      .Detail(Row, "nUnitPrce") = 0#
      .Detail(Row, "sMCInvIDx") = ""
      .Detail(Row, "cMotorNew") = "0"
      .Detail(Row, "sReplMCID") = ""
      .Detail(Row, "nReplAmtx") = 0#
      .Detail(Row, "nReplIAmt") = 0#
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(7)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & Row _
                        & ", " & Value _
                        & ", " & Search & " )"
End Function

Private Function getInsurance(ByVal Row As Integer, _
                           ByVal Value As String, _
                           ByVal Search As Boolean) As Boolean
   Dim lors As Recordset
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "getDetail"
   'On Error Goto errProc
   getInsurance = False
   
   lsSQL = "SELECT" & _
               "  sInsTypID" & _
               ", sInsTypNm" & _
               ", cRegisTyp" & _
            " FROM Insurance_Type" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
   
   If Value <> "" Then
      If Search Then
         lsSQL = AddCondition(lsSQL, "sInsTypNm LIKE " & strParm(Trim(Value) & "%"))
      Else
         lsSQL = AddCondition(lsSQL, "sInsTypNm = " & strParm(Trim(Value)))
      End If
   ElseIf Search = False Then
      GoTo endWithClear
   End If

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, , adCmdText
   
   If lors.EOF Then
      If Search = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   
   With p_oBaseClas
      If lors.RecordCount = 1 Then
         .Detail(Row, "sInsTypID") = lors("sInsTypID")
         .Detail(Row, "sInsTypNm") = lors("sInsTypNm")
         .Detail(Row, "cRegisTyp") = lors("cRegisTyp")
         RaiseEvent DetailRetrieved(5)
         GoTo endProc
      End If
      
      lsSQL = KwikBrowse(p_oAppDrivr, lors, "sInsTypNm", "Insurance Type")
      If lsSQL = "" Then
         If Search Then
            GoTo endProc
         Else
            GoTo endWithClear
         End If
      End If
      lasDetail = Split(lsSQL, "»")
      .Detail(Row, "sInsTypID") = lasDetail(0)
      .Detail(Row, "sInsTypNm") = lasDetail(1)
      .Detail(Row, "cRegisTyp") = lasDetail(2)
   End With
   
   RaiseEvent DetailRetrieved(5)
   
   getInsurance = True
   
endProc:
   Set lors = Nothing
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(Row, "sInsTypID") = ""
      .Detail(Row, "sInsTypNm") = ""
      .Detail(Row, "cRegisTyp") = ""
   End With
   RaiseEvent DetailRetrieved(5)
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Row _
                        & ", " & Value _
                        & ", " & Search & " )"
   GoTo endProc
End Function

Private Sub computeDueDate()
   Dim ldDueDate As Date
   Dim lsOldProc As String
   
   lsOldProc = "computeDueDate"
   'On Error Goto errProc
   
   With p_oMCARAct
      If .Fields("nAcctTerm") = 0 Then
         .Fields("dDueDatex") = .Fields("dFirstPay")
      Else
         .Fields("dDueDatex") = DateAdd("m", .Fields("nAcctTerm") - 1, .Fields("dFirstPay"))
      End If
   End With
   RaiseEvent AccountRetrieved(8)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Sub

Private Sub computeTerm()
   Dim lsOldProc As String
   
   lsOldProc = "computeTerm"
   'On Error Goto errProc
   
   With p_oMCARAct
      .Fields("nGrossPrc") = .Fields("nPNValuex") + .Fields("nDownPaym") + .Fields("nCashBalx")
      .Fields("nABalance") = .Fields("nGrossPrc") - .Fields("nDownTotl")
      If .Fields("nAcctTerm") = 0 Then
         .Fields("nMonAmort") = 0
      Else
         .Fields("nMonAmort") = Round(.Fields("nPNValuex") / .Fields("nAcctTerm"), 2)
      End If

      RaiseEvent AccountRetrieved(9)
      RaiseEvent AccountRetrieved(12)
      RaiseEvent AccountRetrieved(13)
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub computeDefault()
   Dim lsOldProc As String
   
   lsOldProc = "computeTerm"
   'On Error Goto errProc
   
   With p_oMCARAct
      If pbCreatApp Then
         .Fields("nGrossPrc") = p_oBaseClas.Master("nTranTotl") * 1.25
         .Fields("nCashBalx") = 0#
         .Fields("nPNValuex") = .Fields("nGrossPrc") - .Fields("nDownPaym")
         .Fields("nMonAmort") = Round(.Fields("nPNValuex") / .Fields("nAcctTerm"), 2)
         .Fields("nPenaltyx") = 0#
         .Fields("nRebatesx") = 0#
         .Fields("nABalance") = .Fields("nGrossPrc") = .Fields("nDownTotl")
      Else
         .Fields("nGrossPrc") = p_oBaseClas.Master("nTranTotl") - p_oBaseClas.Detail(0, "nInsAmtxx")
         .Fields("nCashBalx") = .Fields("nGrossPrc") - .Fields("nDownPaym")
         .Fields("nPNValuex") = 0#
         .Fields("nMonAmort") = 0#
         .Fields("nPenaltyx") = 0#
         .Fields("nRebatesx") = 0#
         .Fields("nABalance") = .Fields("nGrossPrc") - .Fields("nDownTotl")
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Sub

Private Function UndoTransaction() As Boolean
   Dim loClient As clsClient
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim Row As Long
   
   lsOldProc = "UndoTransaction"
   'On Error Goto errProc
   UndoTransaction = False
   
   Set loClient = New clsClient
   Set loClient.AppDriver = p_oAppDrivr
   
   With p_oBaseClas
      Set lors = New Recordset
      lsSQL = AddCondition(.DetailQuery, _
                  "a.sTransNox = " & strParm(.Master("sTransNox")))
      
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.EOF Then GoTo endProc
      
      If lors.RecordCount = 1 Then
         If unSaveMCUnit(lors) = False Then GoTo endProc
      Else
         If unSaveMCUnits(lors) = False Then GoTo endProc
      End If
      
      ' undo receipt transaction
      If .MasterOriginal("nAmtPaidx") > 0# Then
         lsSQL = "DELETE FROM Receipt_Master" & _
                  " WHERE sORNoxxxx = " & strParm(.MasterOriginal("sORNoxxxx")) & _
                     " AND sClientID = " & strParm(.MasterOriginal("sClientID")) & _
                     " AND dTransact = " & dateParm(.MasterOriginal("dTransact")) & _
                     " AND LEFT(sTransNox, 2) = " & strParm(Left(.MasterOriginal("sTransNox"), 2)) & _
                     " AND sReferNox = " & strParm(.MasterOriginal("sTransNox")) & _
                     " AND nTranTotl = " & .MasterOriginal("nAmtPaidx")

         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

         If Row <= 0 Then
            MsgBox "Unable to Delete Receipt Transaction Info!!!" & lsSQL, vbCritical, "Warning"
            GoTo endProc
         End If

         lsSQL = "DELETE FROM Checks_Received" & _
                  " WHERE sReferNox = " & strParm(p_axReceipt.TransNox)

         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      End If
      
      ' undo client transaction
      If .MasterOriginal("cCustType") = 0 Then
         ' update the client ledger
         loClient.InitTransaction
         loClient.Branch = p_sBranchCd
         loClient.ClientID = .MasterOriginal("sClientID")
         loClient.ClientType = xeClientCustomer
         loClient.DebitAmount = .MasterOriginal("nTranTotl")
         loClient.CreditAmount = .MasterOriginal("nAmtPaidx")
         loClient.SourceSystem = "MC"
         loClient.TransactDate = .MasterOriginal("dTransact")
         If loClient.Sales(.Master("sTransNox"), True) = False Then GoTo endProc
      End If
      
      If .MasterOriginal("sApplicNo") <> Empty Then
         lors.MoveFirst
         If .MasterOriginal("sApplicNo") <> "N-O-N-E" Then
            lsSQL = "SELECT sAcctNmbr FROM MC_AR_Master" & _
                     " WHERE sApplicNo = " & strParm(.MasterOriginal("sApplicNo"))
         Else
            lsSQL = "SELECT sAcctNmbr FROM MC_AR_Master" & _
                     " WHERE sClientID = " & strParm(.MasterOriginal("sClientID")) & _
                        " AND sSerialID = " & strParm(lors("sSerialID")) & _
                        " AND dPurchase = " & dateParm(.MasterOriginal("dTransact"))
         End If

         lors.Close
         lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

         If lors.EOF Then
            MsgBox "Unable to Retrieve LR Account Info!!!" & lsSQL, vbCritical, "Warning"
            GoTo endProc
         End If

         lsSQL = "DELETE FROM MC_AR_Master" & _
                  " WHERE sAcctNmbr = " & strParm(lors("sAcctNmbr"))

         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

         If Row <= 0 Then
            MsgBox "Unable to Delete LR Master Info!!!", vbCritical, "Warning"
            GoTo endProc
         End If

         lsSQL = "DELETE FROM MC_AR_Ledger" & _
                  " WHERE sAcctNmbr = " & strParm(lors("sAcctNmbr"))

         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

         If .MasterOriginal("sApplicNo") <> "N-O-N-E" Then
            lsSQL = "UPDATE MC_Credit_Application SET " & _
                        " cTranStat = " & strParm(xeStatePosted) & _
                     " WHERE sTransNox = " & strParm(.MasterOriginal("sApplicNo"))

            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

            If Row <= 0 Then
               MsgBox "Unable to Update Credit Application Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If

      If undoGiveAways = False Then GoTo endProc
      If undoAccessories = False Then GoTo endProc
   End With
   
   UndoTransaction = True
   
endProc:
   Set loClient = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function SaveInsurance(ByVal lsRefNo As String, _
                               ByVal lnDetail As Integer) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   
   lsOldProc = "SaveInsurance"
   'On Error Goto errProc
   SaveInsurance = False
   
   With p_oBaseClas
      lsSQL = "INSERT INTO MC_Insurance (" & _
                  "  sInsurNox" & _
                  ", sReferNox" & _
                  ", sSourceCd" & _
                  ", sClientID" & _
                  ", sSerialID" & _
                  ", dTransact" & _
                  ", sPolicyNo" & _
                  ", sInsTypID" & _
                  ", cRegisTyp" & _
                  ", nPremiumx" & _
                  ", nAmtPaidX" & _
                  ", cTranStat" & _
                  ", cRecvCOCx" & _
                  ", sModified" & _
                  ", dModified )"
      lsSQL = lsSQL & " VALUES (" & _
                  strParm(lsRefNo) & _
                  ", " & strParm(.Master("sTransNox")) & _
                  ", " & strParm("Sl") & _
                  ", " & strParm(.Master("sClientID")) & _
                  ", " & strParm(.Detail(lnDetail, "sSerialID")) & _
                  ", " & dateParm(.Master("dTransact")) & _
                  ", ''" & _
                  ", " & strParm(.Detail(lnDetail, "sInsTypID")) & _
                  ", " & strParm(.Detail(lnDetail, "cRegisTyp")) & _
                  ", 0.00" & _
                  ", 0.00" & _
                  ", " & xeStateOpen & _
                  ", '0'" & _
                  ", " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                  ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Save MC Registration Expense Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   SaveInsurance = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsRefNo _
                        & ", " & lnDetail & " )"
   GoTo endProc
End Function

Private Function saveExpenses(ByVal lsRefNo As String, _
                              ByVal lnDetail As Integer) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   Dim oLTOPool As clsMCLTOPool
   
   lsOldProc = "SaveExpenses"
   'On Error Goto errProc
   saveExpenses = False
   
   With p_oBaseClas
      lsSQL = "INSERT INTO MC_Registration_Expense (" & _
                  "  sTransNox" & _
                  ", sReferNox" & _
                  ", cMotorNew" & _
                  ", cInsurPrc" & _
                  ", nInsurAmt" & _
                  ", cClearPrc" & _
                  ", nClearAmt" & _
                  ", nClearSOP" & _
                  ", cValidPrc" & _
                  ", nValidAmt" & _
                  ", nValidSOP" & _
                  ", cRegisPrc" & _
                  ", nRegisAmt" & _
                  ", nRegisSOP" & _
                  ", cPlatePrc" & _
                  ", nPlateSOP" & _
                  ", dModified"
      lsSQL = lsSQL & _
               " ) VALUES ( " & _
                  strParm(lsRefNo) & _
                  ", " & strParm(p_sRegisNox) & _
                  ", " & strParm(.Detail(lnDetail, "cMotorNew")) & _
                  ", '', 0.00" & _
                  ", '', 0.00, 0.00" & _
                  ", '', 0.00, 0.00" & _
                  ", '', 0.00, 0.00" & _
                  ", '', 0.00" & _
                  ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Save MC Registration Expense Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Set oLTOPool = New clsMCLTOPool
      Set oLTOPool.AppDriver = p_oAppDrivr
      oLTOPool.TranDate = .Master("dTransact")
      oLTOPool.SerialID = .Detail(lnDetail, "sSerialID")
      Call oLTOPool.SaveCollection( _
           p_oAppDrivr.BranchCode, _
           .Master("sTransNox"), _
           "SRg", _
           .Detail(lnDetail, "nInsAmtxx") _
           )
   
   End With
   
   saveExpenses = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsRefNo _
                        & ", " & lnDetail & " )"
   GoTo endProc
End Function

Private Function SaveDeedSaleExpense(ByVal lsRefNo As String, _
                                       ByVal lnDetail As Integer) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   Dim oLTOPool As clsMCLTOPool
   
   lsOldProc = "SaveDeedSaleExpense"
   'On Error Goto errProc
   SaveDeedSaleExpense = False
   
   With p_oBaseClas
      If .Detail(lnDetail, "cMotorNew") = xeNo Then
         If .Detail(lnDetail, "nInsAmtxx") > 0 Then
            lsSQL = "INSERT INTO MC_DeedSale_Expense (" & _
                        "  sTransNox" & _
                        ", dTransact" & _
                        ", sReferNox" & _
                        ", sSourceCD" & _
                        ", sClientID" & _
                        ", sSerialID" & _
                        ", cDeedSPrc" & _
                        ", nDeedSAmt" & _
                        ", nAmtPaidx" & _
                        ", dModified" & _
                     " ) VALUES ( " & _
                        strParm(lsRefNo) & _
                        ", " & dateParm(.Master("dTransact")) & _
                        ", " & strParm(.Master("sTransNox")) & _
                        ", " & strParm("SDS") & _
                        ", " & strParm(.Master("sClientID")) & _
                        ", " & strParm(.Detail(lnDetail, "sSerialID")) & _
                        ", " & strParm(xeNo) & _
                        ", " & 0# & _
                        ", " & IIf(.Detail(lnDetail, "sInsTypID") = "", _
                           .Detail(lnDetail, "nInsAmtxx"), 0#) & _
                        ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
            
            If Row <= 0 Then
               MsgBox "Unable to Save Deed of Sale Expense Entry!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         
            Set oLTOPool = New clsMCLTOPool
            Set oLTOPool.AppDriver = p_oAppDrivr
            oLTOPool.TranDate = .Master("dTransact")
            oLTOPool.SerialID = .Detail(lnDetail, "sSerialID")
            Call oLTOPool.SaveCollection( _
                 p_oAppDrivr.BranchCode, _
                 .Master("sTransNox"), _
                 "SDS", _
                 IIf(.Detail(lnDetail, "sInsTypID") = "", .Detail(lnDetail, "nInsAmtxx"), 0#) _
                 )
            Set oLTOPool = Nothing
         End If
      End If
   End With

   SaveDeedSaleExpense = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsRefNo _
                        & ", " & lnDetail & " )"
   GoTo endProc
End Function

Private Function SaveRegistration(ByVal lsRefNo As String, _
                                    ByVal lnDetail As Integer) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   
   lsOldProc = "SaveRegistration"
   'On Error Goto errProc
   SaveRegistration = False
    
   With p_oBaseClas
      lsSQL = "INSERT INTO MC_Registration (" & _
                  "  sTransNox" & _
                  ", dPurchase" & _
                  ", sReferNox" & _
                  ", sSourceCD" & _
                  ", sSerialID" & _
                  ", sClientID" & _
                  ", cRegisTyp" & _
                  ", nRegisAmt" & _
                  ", sORNoxxxx" & _
                  ", sRegORNox" & _
                  ", sCRENoxxx" & _
                  ", sCRNoxxxx" & _
                  ", sStickrNo" & _
                  ", sModified" & _
                  ", dModified"
      lsSQL = lsSQL & _
               " ) VALUES (" & _
                  strParm(lsRefNo) & _
                  ", " & dateParm(.Master("dTransact")) & _
                  ", " & strParm(.Master("sTransNox")) & _
                  ", " & strParm("SRg") & _
                  ", " & strParm(.Detail(lnDetail, "sSerialID")) & _
                  ", " & strParm(.Master("sClientID")) & _
                  ", " & strParm(.Detail(lnDetail, "cRegisTyp")) & _
                  ", " & .Detail(lnDetail, "nInsAmtxx") & _
                  ", ''" & _
                  ", ''" & _
                  ", ''" & _
                  ", ''" & _
                  ", ''" & _
                  ", ''" & _
                  ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Save MC Registration Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If .Detail(lnDetail, "cMotorNew") = xeNo Then
         ' repo sales with registration processing fee
         If SaveDeedSaleExpense(GetNextCode("MC_DeedSale_Expense", "sTransNox", True, _
                           p_oAppDrivr.Connection, True, p_sBranchCd), lnDetail) = False Then GoTo endProc
      End If
   End With
   
   SaveRegistration = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsRefNo _
                        & ", " & lnDetail & " )"
   GoTo endProc
End Function

Private Function UpdateRegistration(ByVal lsRefNo As String, _
                                  ByVal lnDetail As Integer) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   Dim lorsx As ADODB.Recordset
   Dim oLTOPool As clsMCLTOPool
   
   lsOldProc = "UpdateRegistration"
   'On Error Goto errProc
   UpdateRegistration = False
   
   Set lorsx = New ADODB.Recordset
   
   With p_oBaseClas
      lsSQL = " SELECT sTransNox" & _
              " FROM MC_Registration" & _
              " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                " AND sSourceCD = " & strParm("SRg") & _
                " AND sSerialID = " & strParm(.Detail(lnDetail, "sReplMCID"))
     lorsx.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockOptimistic, adCmdText
      If lorsx.EOF Then
         MsgBox "Unable to Delete MC Registration Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "UPDATE MC_Registration SET" & _
                  "  sSerialID = " & strParm(.Detail(lnDetail, "sSerialID")) & _
                  ", sReferNox = " & strParm(.Master("sTransNox")) & _
                  ", dPurchase = " & dateParm(.Master("dTransact")) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(lorsx("sTransNox"))
               
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Update MC Registration Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "UPDATE MC_Registration_Expense SET" & _
                  "  sSerialID = " & strParm(.Detail(lnDetail, "sSerialID")) & _
                  ", dTransact = " & dateParm(.Master("dTransact")) & _
                  ", cMotorNew = " & strParm(.Detail(lnDetail, "cMotorNew")) & _
                  ", cInsurPrc = " & strParm(xeNo) & _
                  ", nInsurAmt = " & 0# & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sReferNox = " & strParm(lorsx("sTransNox"))
      
      
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Update MC Registration Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Set oLTOPool = New clsMCLTOPool
      Set oLTOPool.AppDriver = p_oAppDrivr
      oLTOPool.TranDate = .Master("dTransact")
      oLTOPool.SerialID = .Detail(lnDetail, "sSerialID")
      Call oLTOPool.SaveCollection( _
           p_oAppDrivr.BranchCode, _
           Master("sTransNox"), _
           "SRg", _
           0# _
           )
      
      lsSQL = "SELECT cTranStat FROM MC_Insurance" & _
              " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                " AND sSourceCD = " & strParm("Sl") & _
                " AND sSerialID = " & strParm(.Detail(lnDetail, "sSerialID"))
      
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then
         MsgBox "Unable to Retrieve Insurance Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If lors("cTranStat") <> xeStateOpen Then
         lsSQL = "UPDATE MC_Insurance SET" & _
                     "  cTranStat = " & strParm(xeStateCancelled) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                 " AND sSourceCD = " & strParm("Sl") & _
                 " AND sSerialID = " & strParm(.Detail(lnDetail, "sSerialID"))

         
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
         
         If Row <= 0 Then
            MsgBox "Unable to Update Insurance Entry!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      
         If SaveInsurance(lsRefNo, lnDetail) = False Then GoTo endProc
      Else
         lsSQL = "UPDATE MC_Insurance SET" & _
                     "  sSerialID = " & strParm(.Detail(lnDetail, "sSerialID")) & _
                     ", dTransact = " & dateParm(.Master("dTransact")) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                  " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                    " AND sSourceCD = " & strParm("Sl") & _
                    " AND sSerialID = " & strParm(.Detail(lnDetail, "sSerialID"))
                    

         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
         
         If Row <= 0 Then
            MsgBox "Unable to Update Insurance Entry!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With
   
   UpdateRegistration = True

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsRefNo _
                        & ", " & lnDetail & " )"
   GoTo endProc
End Function

Private Function UndoUpdateRegistration(ByVal lsSerialID As String, _
                                        ByVal lsReplMCID As String) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lsInsTypID As String, lcRegisTyp As String
   Dim Row As Long
   Dim lorsx As ADODB.Recordset
   Dim lsTransNox As String
   Dim oLTOPool As clsMCLTOPool
   
   lsOldProc = "UndoUpdateRegistration"
   'On Error Goto errProc
   UndoUpdateRegistration = False
   
   Set lorsx = New ADODB.Recordset
   
   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  a.sTransNox" & _
                  ", a.dTransact" & _
                  ", b.cMotorNew" & _
                  ", b.sInsTypID" & _
                  ", b.cRegisTyp" & _
               " FROM MC_SO_Master a" & _
                  ", MC_SO_Detail b" & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND a.sClientID = " & strParm(.MasterOriginal("sClientID")) & _
                  " AND b.sSerialID = " & strParm(lsReplMCID)
      
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then
         MsgBox "Unable to Retrieve Registration Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsInsTypID = lors("sInsTypID")
      lcRegisTyp = lors("cRegisTyp")
      lsTransNox = lors("sTransNox")
      
      lsSQL = " SELECT sTransNox" & _
              " FROM MC_Registration" & _
              " WHERE sReferNox = " & strParm(lsTransNox) & _
                " AND sSourceCD = " & strParm("SRg") & _
                " AND sSerialID = " & strParm(lsSerialID)
                
      lorsx.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockOptimistic, adCmdText
      If lorsx.EOF Then
         MsgBox "Unable to retriever MC Registration Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "UPDATE MC_Registration SET" & _
                  "  sSerialID = " & strParm(lsReplMCID) & _
                  ", sTransNox = " & strParm(lors("sTransNox")) & _
                  ", dPurchase = " & dateParm(lors("dTransact")) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sTransNox = " & strParm(lorsx("sTransNox"))
                  
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Update MC Registration Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "UPDATE MC_Registration_Expense SET" & _
                  ", cMotorNew = " & strParm(lors("cMotorNew")) & _
                  ", cInsurPrc = " & strParm(xeNo) & _
                  ", nInsurAmt = " & 0# & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sReferNox = " & strParm(lorsx("sTransNox"))
      
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Update MC Registration Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "SELECT * FROM MC_Insurance" & _
              " WHERE sReferNox = " & strParm(lsTransNox) & _
                " AND sSourceCD = " & strParm("Sl") & _
                " AND sSerialID = " & strParm(lsReplMCID)
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then
         lsSQL = "UPDATE MC_Insurance SET" & _
                     "  sSerialID = " & strParm(lsReplMCID) & _
                     ", dTransact = " & dateParm(.MasterOriginal("dTransact")) & _
                     ", cRegisTyp = " & strParm(lcRegisTyp) & _
                     ", sInsTypID = " & strParm(lsInsTypID) & _
                     ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                 " WHERE sReferNox = " & strParm(lsTransNox) & _
                   " AND sSourceCD = " & strParm("Sl") & _
                   " AND sSerialID = " & strParm(lsSerialID)
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
         
         If Row <= 0 Then
            MsgBox "Unable to Update Insurance Entry!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      Else
         If lors("cTranStat") = xeStateCancelled Then
            lsSQL = "UPDATE MC_Insurance SET" & _
                        "  cTranStat = " & strParm(xeStateOpen) & _
                        ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
                    " WHERE sReferNox = " & strParm(lsTransNox) & _
                      " AND sSourceCD = " & strParm("Sl") & _
                      " AND sSerialID = " & strParm(lsReplMCID)
            
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
            
            If Row <= 0 Then
               MsgBox "Unable to Update Insurance Entry!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         
            lsSQL = "DELETE FROM MC_Insurance" & _
                    " WHERE sReferNox = " & strParm(lsTransNox) & _
                      " AND sSourceCD = " & strParm("Sl") & _
                      " AND sSerialID = " & strParm(lsSerialID)
            
            p_oAppDrivr.RegisDelete lsSQL
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
            
            If Row <= 0 Then
               MsgBox "Unable to Delete Customer Insurance Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         
            Set oLTOPool = New clsMCLTOPool
            Set oLTOPool.AppDriver = p_oAppDrivr
            oLTOPool.TranDate = .Master("dTransact")
            oLTOPool.SerialID = lsSerialID
            Call oLTOPool.SaveCollection( _
                 p_oAppDrivr.BranchCode, _
                 Master("sTransNox"), _
                 "SRg", _
                 0# _
                 )
         
         End If
      End If
   End With
   
   UndoUpdateRegistration = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsSerialID _
                        & ", " & lsReplMCID & " )"
   GoTo endProc
End Function

Private Function SaveReceipt() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lsRefNo As String
   Dim lnRefNo As Long
   Dim Row As Long
   Dim lnCtr As Integer
   
   lsOldProc = "SaveReceipt"
   'On Error Goto errProc
   SaveReceipt = False
   
   With p_oBaseClas
      p_axReceipt.TransNox = GetNextCode("Receipt_Master", "sTransNox", True, p_oAppDrivr.Connection, True, p_sBranchCd)
      lsSQL = "INSERT INTO Receipt_Master (" & _
                  "  sTransNox" & _
                  ", sORNoxxxx" & _
                  ", dTransact" & _
                  ", sSystemCd" & _
                  ", sClientID" & _
                  ", nCashAmtx" & _
                  ", nCheckAmt" & _
                  ", nDiscount" & _
                  ", nOthersxx" & _
                  ", cOffCollt" & _
                  ", nTranTotl" & _
                  ", sRemarksx" & _
                  ", cTranType" & _
                  ", sAcctNmbr" & _
                  ", sSerialID" & _
                  ", sInsTypID" & _
                  ", cCashierx" & _
                  ", sAcctCode" & _
                  ", sReferNox" & _
                  ", cTranStat" & _
                  ", sApproved" & _
                  ", sModified" & _
                  ", dModified"
      lsSQL = lsSQL & " ) VALUES (" & _
                  " " & strParm(p_axReceipt.TransNox) & _
                  ", " & strParm(p_axReceipt.ORNoxxxx) & _
                  ", " & dateParm(.Master("dTransact")) & _
                  ", " & strParm("MC") & _
                  ", " & strParm(.Master("sClientID")) & _
                  ", " & p_axReceipt.CashAmtx & _
                  ", " & p_axReceipt.CheckAmt & _
                  ", 0.00" & _
                  ", 0.00" & _
                  ", " & strParm(xeYes) & _
                  ", " & p_axReceipt.TranTotl & _
                  ", " & strParm(p_axReceipt.Remarksx) & _
                  ", " & strParm(0) & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("1") & _
                  ", " & strParm("") & _
                  ", " & strParm(.Master("sTransNox")) & _
                  ", " & strParm(xeStateOpen) & _
                  ", " & strParm("") & _
                  ", " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                  ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Update Transaction Payment!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If p_axReceipt.CheckAmt > 0 Then
         lsRefNo = GetNextCode("Checks_Received", "sTransNox", True, _
                     p_oAppDrivr.Connection, True, p_sBranchCd)
         lnRefNo = CLng(Mid(lsRefNo, 5))
         For lnCtr = 0 To 4
            If p_axChecks(lnCtr).Amountxx > 0# Then
               lsSQL = "INSERT INTO Checks_Received (" & _
                           "  sTransNox" & _
                           ", sReferNox" & _
                           ", sCheckNox" & _
                           ", sAcctNoxx" & _
                           ", sBankName" & _
                           ", dCheckDte" & _
                           ", nAmountxx" & _
                           ", nClearing" & _
                           ", cChckStat" & _
                           ", cDepositd" & _
                           ", dModified" & _
                        " ) VALUES ( "
               lsSQL = lsSQL & _
                           " " & strParm(lsRefNo) & _
                           ", " & strParm(p_axReceipt.TransNox) & _
                           ", " & strParm(p_axChecks(lnCtr).CheckNox) & _
                           ", " & strParm(p_axChecks(lnCtr).AcctNoxx) & _
                           ", " & strParm(p_axChecks(lnCtr).BankName) & _
                           ", " & dateParm(p_axChecks(lnCtr).CheckDte) & _
                           ", " & p_axChecks(lnCtr).Amountxx & _
                           ", 7" & _
                           ", 0" & _
                           ", 0" & _
                           ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
               p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
               If Row <= 0 Then
                  MsgBox "Unable to Update Check Payment!!!", vbCritical, "Warning"
                  GoTo endProc
               End If
               
               lnRefNo = lnRefNo + 1
               lsRefNo = p_sBranchCd & Format(p_oAppDrivr.ServerDate, "YY") & _
                           Format(lnRefNo, String(6, "0"))
            End If
         Next
      End If
   End With
   
   SaveReceipt = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function saveGiveAways() As Boolean
   Dim loSpareparts As clsSpareparts
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim Row As Long
   Dim lbGiven As Boolean
   
   lsOldProc = "saveGiveAways"
   'On Error Goto errProc
   saveGiveAways = False
   
   Set lors = New Recordset
   Set loSpareparts = New clsSpareparts
   Set loSpareparts.AppDriver = p_oAppDrivr
   
   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  c.sPartsIDx" & _
                  ", SUM(c.nQuantity) xQuantity" & _
                  ", d.nQtyOnHnd" & _
                  ", e.sTransNox" & _
               " FROM MC_SO_Detail a" & _
                     " LEFT JOIN MC_SO_Detail e" & _
                        " ON a.sTransNox = e.sTransNox" & _
                           " AND a.nEntryNox = e.nEntryNox" & _
                           " AND e.cMotorNew = " & strParm(1) & _
                  ", MC_Serial b" & _
                  ", MC_Sales_GiveAways c" & _
                  ", SP_Inventory d" & _
               " WHERE a.sSerialID = b.sSerialID" & _
                  " AND b.sModelIDx = c.sModelIDx" & _
                  " AND c.sPartsIDx = d.sPartsIDx" & _
                  " AND c.sBranchCd = d.sBranchCd" & _
                  " AND a.sTransNox = " & strParm(.Master("sTransNox")) & _
                  " AND d.sBranchCd = " & strParm(p_sBranchCd) & _
               " GROUP BY c.sPartsIDx" & _
                  ", c.sModelIDx" & _
                  ", d.nQtyOnHnd" & _
                  ", e.sTransNox"
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then
         MsgBox "No GiveAway is Assigned to the Unit Sold!!!" & vbCrLf & _
                  "Please verify MC Sales GiveAway and Act Accordingly!!!" & vbCrLf & vbCrLf & _
                  "No GiveAway Transaction was Created at This Time!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      p_oAppDrivr.Connection.BeginTrans
      
      loSpareparts.InitTransaction
      loSpareparts.Branch = p_sBranchCd
      loSpareparts.TransactDate = .Master("dTransact")
      
      ' check first if a giveaway must be given...
      lbGiven = False
      Do
         If IsNull(lors("sTransNox")) = False Then
            lbGiven = True
            Exit Do
         End If
         lors.MoveNext
      Loop
      
      lors.MoveFirst
      If lbGiven = False Then GoTo endProc
      
      pnCtr = 1
      Do
         Row = IIf(lors("nQtyOnHnd") < lors("xQuantity"), lors("xQuantity") - lors("nQtyOnHnd"), 0)
         If Row > 0 Then lbGiven = False
         
         loSpareparts.PartsID = lors("sPartsIDx")
         loSpareparts.Quantity = lors("xQuantity") - Row
         loSpareparts.TransactDate = .Master("dTransact")
         
         If loSpareparts.SalesGiveAway(.Master("sTransNox")) = False Then
            p_oAppDrivr.Connection.RollbackTrans
            GoTo endProc
         End If
         
         lsSQL = "INSERT INTO MC_SO_GiveAways " & _
                     "( sTransNox" & _
                     ", nEntryNox" & _
                     ", sPartsIDx" & _
                     ", nQuantity" & _
                     ", nGivenxxx" & _
                     ", dModified" & _
                  " ) VALUES ( " & _
                     strParm(.Master("sTransNox")) & _
                     ", " & pnCtr & _
                     ", " & strParm(lors("sPartsIDx")) & _
                     ", " & lors("xQuantity") & _
                     ", " & lors("xQuantity") - Row & _
                     ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
         p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
         
         If Row <= 0 Then
            MsgBox "Unable to Save MC Sales GiveAways Transaction!!!" & vbCrLf & vbCrLf & _
                     "Please Make Sure to Create GiveAway Transaction Later!!!", vbCritical, "Warning"
            p_oAppDrivr.Connection.RollbackTrans
            GoTo endProc
         End If
         
         pnCtr = pnCtr + 1
         lors.MoveNext
      Loop Until lors.EOF
      p_oAppDrivr.Connection.CommitTrans
   End With
   
   If lbGiven = False Then MsgBox "Some GiveAways Included in the Transaction were not Given!!!", vbCritical, "Warning"
   
   saveGiveAways = True
   
endProc:
   Set loSpareparts = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

'  use to save a single mc unit purchase
Private Function saveMCUnit() As Boolean
   Dim loMCTrans As clsMCSerialTrans
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lasRefNo(2) As String
   Dim lsOldProc As String
   Dim Row As Long
   
   lsOldProc = "SaveMCUnit"
   'On Error Goto errProc
   saveMCUnit = False
   
   With p_oBaseClas
      If .Detail(0, "sEngineNo") = "" And .Detail(0, "sRemarksx") <> Empty Then
         saveMCUnit = True
         GoTo endProc
      End If
      
      ' validate motorcycle registration, then show warning if motorcycle has no
      '  valid registration
      If .Detail(0, 5) = Empty And .Detail(0, "cMotorNew") = 1 Then
         Row = MsgBox("Motorcycle Sales Does Not Have Registration!!!" & vbCrLf & _
                        "    Continue Saving Anyway?", vbQuestion + vbYesNo, "Confirm")
         If Row <> vbYes Then GoTo endProc
      ElseIf .Detail(0, "cMotorNew") = xeNo And .Detail(0, "nInsAmtxx") = 0 Then
         Row = MsgBox("Motorcycle Sales Does Not Have Deed of Sale Payment!!!" & vbCrLf & _
                        "    Continue Saving Anyway?", vbQuestion + vbYesNo, "Confirm")
         If Row <> vbYes Then GoTo endProc
      End If
      
      ' insert a registration record for the client
      lasRefNo(0) = GetNextCode("MC_Registration", "sTransNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)
      p_sRegisNox = lasRefNo(0)
      
      lasRefNo(1) = GetNextCode("MC_Registration_Expense", "sTransNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)
      lasRefNo(2) = GetNextCode("MC_Insurance", "sInsurNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)
      
      Set loMCTrans = New clsMCSerialTrans
      Set loMCTrans.AppDriver = p_oAppDrivr
      loMCTrans.Branch = p_sBranchCd
      
      loMCTrans.InitTransaction
      loMCTrans.TransactDate = .Master("dTransact")
      loMCTrans.MCInventoryID = .Detail(0, "sMCInvIDx")
      loMCTrans.Quantity = 1
      
      loMCTrans.Serial(0, "SerialID") = .Detail(0, "sSerialID")
      loMCTrans.Serial(0, "EngineNo") = .Detail(0, "sEngineNo")
      loMCTrans.Serial(0, "SerialID") = .Detail(0, "sSerialID")
      loMCTrans.Serial(0, "WarrntNo") = .Detail(0, "sWarrntNo")
            
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "nEntryNox") = 1
            
      ' insert registration only if registration payment is detected
      If .Detail(0, 5) <> Empty Then
         saveMCUnit = SaveRegistration(lasRefNo(0), 0) And _
                      saveExpenses(lasRefNo(1), 0) And _
                      SaveInsurance(lasRefNo(2), 0)

         If saveMCUnit = False Then GoTo endProc
      ElseIf .Detail(0, 6) > 0 Then
         lasRefNo(0) = GetNextCode("MC_DeedSale_Expense", "sTransNox", True, _
                           p_oAppDrivr.Connection, True, p_sBranchCd)
         saveMCUnit = SaveDeedSaleExpense(lasRefNo(0), 0)
         
         If saveMCUnit = False Then GoTo endProc
      End If
      
      ' if a serial is not successfully updated, cancel saving.
      saveMCUnit = loMCTrans.Sales(.Master("sTransNox"), False, .Master("sClientID"))
      If saveMCUnit = False Then GoTo endProc
      
      ' now check if the sold unit is a replacement
      If .Master("cTranType") = 2 Then
         loMCTrans.MCInventoryID = .Detail(0, "sRMCInvID")
         loMCTrans.Quantity = 1
         
         loMCTrans.Serial(0, "SerialID") = .Detail(0, "sReplMCID")
         If loMCTrans.SalesReturn(.Master("sTransNox")) = False Then GoTo endProc
         
         ' now check if no registration payment was made...
         If .Detail(0, "sInsTypID") = Empty Then
            ' now check if previous sales has registration...
            If .Detail(0, "nReplIAmt") > 0 Then
               saveMCUnit = UpdateRegistration(lasRefNo(2), 0)
               If saveMCUnit = False Then GoTo endProc
            End If
         End If
      End If
      
      ' update the impounded account of motorcycle sold...
      If .Detail(0, "cMotorNew") = xeNo Then
         If UpdateImpounded(.Detail(0, "sSerialID")) = False Then GoTo endProc
      End If
   End With
   
   saveMCUnit = True
   
endProc:
   Set loMCTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function unSaveMCUnit(lors As Recordset) As Boolean
   Dim loMCTrans As clsMCSerialTrans
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim Row As Long
   Dim lorsx As ADODB.Recordset
   
   lsOldProc = "unSaveMCUnit"
   'On Error Goto errProc
   unSaveMCUnit = False
   
   Set loMCTrans = New clsMCSerialTrans
   Set loMCTrans.AppDriver = p_oAppDrivr
   Set lorsx = New ADODB.Recordset
   
   loMCTrans.Branch = p_sBranchCd
   loMCTrans.InitTransaction
   
   With p_oBaseClas
      If lors("sEngineNo") = "" And lors("sRemarksx") <> Empty Then
         unSaveMCUnit = True
         GoTo endProc
      End If
      
      ' undo motorcycle transaction
      loMCTrans.MCInventoryID = lors("sMCInvIDx")
      loMCTrans.Quantity = 1
      loMCTrans.Serial(0, "SerialID") = lors("sSerialID")
      
      If loMCTrans.Sales(.Master("sTransNox"), True) = False Then GoTo endProc
   
      ' now check if the sold unit is a replacement
      If .MasterOriginal("cTranType") = 2 Then
         loMCTrans.MCInventoryID = lors("sRMCInvID")
         loMCTrans.Quantity = 1
         
         loMCTrans.Serial(0, "SerialID") = lors("sReplMCID")
         If loMCTrans.SalesReturn(.Master("sTransNox"), True) = False Then GoTo endProc
      
         ' now check if no registration payment was made...
         If lors("sInsTypID") = Empty Then
            ' now check if previous sales has registration...
            If lors("nReplIAmt") > 0 Then
               unSaveMCUnit = UndoUpdateRegistration(lors("sSerialID"), lors("sReplMCID"))
               GoTo endProc
            End If
         End If
      Else
         If lors("sInsTypID") <> Empty Then
            
            lsSQL = " SELECT sTransNox" & _
                    " FROM MC_Registration" & _
                    " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                      " AND sSourceCD = " & strParm("SRg") & _
                      " AND sSerialID = " & strParm(lors("sSerialID"))
                      
            lorsx.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockOptimistic, adCmdText
            If lorsx.EOF Then
               MsgBox "Unable to retriever MC Registration Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If
            
            ' undo registration transaction
            lsSQL = "DELETE FROM MC_Registration" & _
                     " WHERE sTransNox = " & strParm(lorsx("sTransNox"))
                     
            p_oAppDrivr.RegisDelete lsSQL
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

            If Row <= 0 Then
               MsgBox "Unable to Delete MC Registration Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If

            ' after undoing registration transaction, process its expenses
            lsSQL = "DELETE FROM MC_Registration_Expense" & _
                     " WHERE sReferNox = " & strParm(lorsx("sTransNox"))

            p_oAppDrivr.RegisDelete lsSQL
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

            If Row <= 0 Then
               MsgBox "Unable to Delete MC Registration Expense Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If

            ' lastly, undo insurance transaction
            lsSQL = "DELETE FROM MC_Insurance" & _
                     " WHERE sSerialID = " & strParm(lors("sSerialID")) & _
                       " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                       " AND sSourceCD = " & strParm("Sl")

            p_oAppDrivr.RegisDelete lsSQL
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText

            If Row <= 0 Then
               MsgBox "Unable to Delete Customer Insurance Info!!!", vbCritical, "Warning"
               GoTo endProc
            End If

            If lors("cMotorNew") = xeNo Then
               ' if a deed of sale expense is created then remove it
               lsSQL = "DELETE FROM MC_DeedSale_Expense" & _
                      " WHERE sSerialID = " & strParm(lors("sSerialID")) & _
                        " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                        " AND sSourceCD = " & strParm("SDS")

               p_oAppDrivr.RegisDelete lsSQL
               p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
            End If
         ElseIf lors("nInsAmtxx") > 0 And lors("cMotorNew") = xeNo Then
            lsSQL = "DELETE FROM MC_DeedSale_Expense" & _
                      " WHERE sSerialID = " & strParm(lors("sSerialID")) & _
                        " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                        " AND sSourceCD = " & strParm("SDS")
            
            p_oAppDrivr.RegisDelete lsSQL
            p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
         End If
      End If
      
      If lors("cMotorNew") = xeNo Then
         If UpdateImpounded(lors("sSerialID"), True) = False Then GoTo endProc
      End If
   End With

   unSaveMCUnit = True
   
endProc:
   Set loMCTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & "lors" & " )"
   GoTo endProc
End Function

'  Update impounded account
Private Function UpdateImpounded(lsSerialID As String, Optional lbReverse As Variant = False) As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String, lsSQL As String
   Dim Row As Long
   
   lsOldProc = "unSaveMCUnit"
   'On Error Goto errProc
   
   lsSQL = "SELECT" & _
               "  sAcctNmbr" & _
               ", sClientID" & _
            " FROM MC_AR_Master" & _
            " WHERE sSerialID = " & strParm(lsSerialID) & _
               " AND cActivexx = " & strParm(xeYes)
               
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then
      ' allow it to update nothing
      UpdateImpounded = True
      GoTo endProc
   End If
   
   With p_oAppDrivr
      lsSQL = "UPDATE MC_AR_Master SET" & _
                  "  cActivexx = " & IIf(lbReverse, strParm(xeYes), strParm(xeNo)) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sAcctNmbr = " & strParm(lors("sAcctNmbr"))
      
      .Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Update Former Owner of Motorcycle!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lsSQL = "UPDATE Impound SET" & _
                  "  cTranStat = " & IIf(lbReverse, strParm(0), strParm(2)) & _
                  ", dModified = " & dateParm(p_oAppDrivr.ServerDate()) & _
               " WHERE sAcctNmbr = " & strParm(lors("sAcctNmbr"))

      If Row <= 0 Then
         MsgBox "Unable to Update Impounded Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   UpdateImpounded = True
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & lsSerialID _
                        & ", " & lbReverse & " )"
   GoTo endProc
End Function

'  use to save mc units purchase
Private Function saveMCUnits() As Boolean
   Dim loMCTrans As clsMCSerialTrans
   Dim lsMCInvID As String, lsOldProc As String
   Dim lasRefNo(2) As String, lasMCInv() As String
   Dim lanRefNo(2) As Long, lanMCInv() As Integer
   Dim Row As Long, lnCtr As Integer
   Dim lnEntryNo As Integer, lnInsurance As Integer
   Dim lbCancel As Boolean
   
   lsOldProc = "SaveMCUnits"
   'On Error Goto errProc
   saveMCUnits = True
   
   With p_oBaseClas
      Set loMCTrans = New clsMCSerialTrans
      Set loMCTrans.AppDriver = p_oAppDrivr
      loMCTrans.Branch = p_sBranchCd
      
      ' insert a registration record for the client
      lasRefNo(0) = GetNextCode("MC_Registration", "sTransNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)
      lanRefNo(0) = CLng(Mid(lasRefNo(0), 5))
      
      lasRefNo(1) = GetNextCode("MC_Registration_Expense", "sTransNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)
      lanRefNo(1) = CLng(Mid(lasRefNo(1), 5))
      
      lasRefNo(2) = GetNextCode("MC_Insurance", "sInsurNox", True, _
                        p_oAppDrivr.Connection, True, p_sBranchCd)
      lanRefNo(2) = CLng(Mid(lasRefNo(2), 5))
      
      lsMCInvID = Empty
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sEngineNo") = "" Then Exit For
         
         If Not IsNull(.Detail(pnCtr, "sMCInvIDx")) Then
            lnCtr = InStr(1, lsMCInvID, .Detail(pnCtr, "sMCInvIDx"))
            If lnCtr = 0 Then
               If lsMCInvID = Empty Then
                  lnCtr = 0
               Else
                  lnCtr = UBound(lasMCInv) + 1
               End If
               
               ReDim Preserve lasMCInv(lnCtr) As String
               ReDim Preserve lanMCInv(lnCtr) As Integer
               
               lasMCInv(lnCtr) = .Detail(pnCtr, "sMCInvIDx")
               lsMCInvID = lsMCInvID & .Detail(pnCtr, "sMCInvIDx") & "»"
            Else
               lnCtr = lnCtr \ 7
               If lnCtr > 6 Then lnCtr = lnCtr - 1
            End If
            lanMCInv(lnCtr) = lanMCInv(lnCtr) + 1
         End If
      Next
      
      lnEntryNo = 1
      For pnCtr = 0 To UBound(lanMCInv)
         loMCTrans.InitTransaction
         loMCTrans.TransactDate = .Master("dTransact")
         loMCTrans.MCInventoryID = lasMCInv(pnCtr)
         loMCTrans.Quantity = lanMCInv(pnCtr)
         
         Row = 0
         lnInsurance = vbCancel
         
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sMCInvIDx") = lasMCInv(pnCtr) Then
               loMCTrans.Serial(Row, "SerialID") = .Detail(lnCtr, "sSerialID")
               loMCTrans.Serial(Row, "EngineNo") = .Detail(lnCtr, "sEngineNo")
               loMCTrans.Serial(Row, "SerialID") = .Detail(lnCtr, "sSerialID")
               loMCTrans.Serial(Row, "WarrntNo") = .Detail(lnCtr, "sWarrntNo")
               
               .Detail(lnCtr, "sTransNox") = .Master("sTransNox")
               .Detail(lnCtr, "nEntryNox") = lnEntryNo
               Row = Row + 1
               lnEntryNo = lnEntryNo + 1
                           
               ' validate motorcycle registration and/or deed of sale
               If lnInsurance = vbCancel Then
                  If .Detail(lnCtr, 5) = Empty And .Detail(lnCtr, "cMotorNew") = 1 Then
                     Row = MsgBox("Some of the Motorcycles Sold Does Not Have Registration!!!" & vbCrLf & _
                                    "    Continue Saving Anyway?", vbQuestion + vbYesNo, "Confirm")
                     If Row <> vbYes Then
                        GoTo endProc
                     Else
                        lnInsurance = vbYes
                     End If
                  ElseIf .Detail(0, "cMotorNew") = xeNo And .Detail(0, "nInsAmtxx") = 0 Then
                     Row = MsgBox("Some of the motorcycle Sold Does Not Have Deed of Sale Payment!!!" & vbCrLf & _
                                    "    Continue Saving Anyway?", vbQuestion + vbYesNo, "Confirm")
                     If Row <> vbYes Then
                        GoTo endProc
                     Else
                        lnInsurance = vbYes
                     End If
                  End If
               End If
               ' insert registration only if registration payment is detected
               If .Detail(lnCtr, 5) <> Empty Then
                  p_sRegisNox = lasRefNo(0)

                  lbCancel = Not (SaveRegistration(lasRefNo(0), lnCtr) And _
                                 saveExpenses(lasRefNo(1), lnCtr) And _
                                 SaveInsurance(lasRefNo(2), lnCtr))
                  
                  If lbCancel Then GoTo endProc
                  
                  lanRefNo(0) = lanRefNo(0) + 1
                  lasRefNo(0) = p_sBranchCd & Format(p_oAppDrivr.ServerDate, "YY") & _
                                    Format(lanRefNo(0), String(6, "0"))
                  
                  lanRefNo(1) = lanRefNo(1) + 1
                  lasRefNo(1) = p_sBranchCd & Format(p_oAppDrivr.ServerDate, "YY") & _
                                    Format(lanRefNo(1), String(6, "0"))
                  
                  lanRefNo(2) = lanRefNo(2) + 1
                  lasRefNo(2) = p_sBranchCd & Format(p_oAppDrivr.ServerDate, "YY") & _
                                    Format(lanRefNo(2), String(6, "0"))
               ElseIf .Detail(lnCtr, 6) > 0 Then
                  lasRefNo(0) = GetNextCode("MC_DeedSale_Expense", "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
                  lbCancel = Not SaveDeedSaleExpense(lasRefNo(0), lnCtr)
                  
                  If lbCancel Then GoTo endProc
               End If
               
               If UpdateImpounded(.Detail(lnCtr, "sSerialID")) = False Then GoTo endProc
            End If
         Next
    
         ' if a serial is not successfully updated, cancel saving.
         saveMCUnits = loMCTrans.Sales(.Master("sTransNox"), False, .Master("sClientID"))
         If saveMCUnits = False Then GoTo endProc
      Next
   End With
   
   saveMCUnits = True
   
endProc:
   Set loMCTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

'  use to undo mc units purchase
Private Function unSaveMCUnits(lors As Recordset) As Boolean
   Dim loMCTrans As clsMCSerialTrans
   Dim lsOldProc As String
   Dim lsMCInvID As String
   Dim lasMCInv() As String
   Dim lanMCInv() As Integer
   Dim lnCtr As Integer, lnSerial As Integer
   Dim lsSQL As String
   Dim Row As Long
   Dim lorsx As ADODB.Recordset
   
   lsOldProc = "unSaveMCUnits"
   'On Error Goto errProc
   unSaveMCUnits = False
   
   Set loMCTrans = New clsMCSerialTrans
   Set loMCTrans.AppDriver = p_oAppDrivr
   Set lorsx = New ADODB.Recordset
   
   loMCTrans.Branch = p_sBranchCd
   loMCTrans.InitTransaction
   
   With p_oBaseClas
      lsMCInvID = Empty
      Do
         If Not IsNull(lors("sMCInvIDx")) Then
            lnCtr = InStr(1, lsMCInvID, lors("sMCInvIDx"))
            If lnCtr = 0 Then
               If lsMCInvID = Empty Then
                  lnCtr = 0
               Else
                  lnCtr = UBound(lasMCInv) + 1
               End If
               
               ReDim Preserve lasMCInv(lnCtr) As String
               ReDim Preserve lanMCInv(lnCtr) As Integer
               
               lasMCInv(lnCtr) = lors("sMCInvIDx")
               lsMCInvID = lsMCInvID & lors("sMCInvIDx") & "»"
            Else
               lnCtr = lnCtr \ 7
               If lnCtr > 6 Then lnCtr = lnCtr - 1
            End If
            lanMCInv(lnCtr) = lanMCInv(lnCtr) + 1
         End If
         
         lors.MoveNext
      Loop Until lors.EOF
      
      For lnCtr = 0 To UBound(lanMCInv)
         loMCTrans.InitTransaction
         loMCTrans.TransactDate = .MasterOriginal("dTransact")
         loMCTrans.MCInventoryID = lasMCInv(lnCtr)
         loMCTrans.Quantity = lanMCInv(lnCtr)
         
         lnSerial = 0
         
         lors.MoveFirst
         Do
            If lors("sMCInvIDx") = lasMCInv(lnCtr) Then
               loMCTrans.Serial(lnSerial, "SerialID") = lors("sSerialID")
               loMCTrans.Serial(lnSerial, "EngineNo") = lors("sEngineNo")
               loMCTrans.Serial(lnSerial, "SerialID") = lors("sSerialID")
               loMCTrans.Serial(lnSerial, "WarrntNo") = lors("sWarrntNo")
               
               lnSerial = lnSerial + 1
            
               If .MasterOriginal("cTranType") <> 2 Then
                  If lors("sInsTypID") <> Empty Then
                     'Get the Registration Number
                     lsSQL = " SELECT sTransNox" & _
                             " FROM MC_Registration" & _
                             " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                               " AND sSourceCD = " & strParm("SRg") & _
                               " AND sSerialID = " & strParm(lors("sSerialID"))
                     
                     lorsx.Open lsSQL, p_oAppDrivr.Connection, adOpenForwardOnly, adLockOptimistic, adCmdText
                     If lorsx.EOF Then
                        MsgBox "Unable to Delete MC Registration Info!!!", vbCritical, "Warning"
                        GoTo endProc
                     End If
                     
                     ' undo registration transaction
                     lsSQL = "DELETE FROM MC_Registration" & _
                              " WHERE sTransNox = " & strParm(lorsx("sTransNox"))
                     p_oAppDrivr.RegisDelete lsSQL
                     p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
                     
                     If Row <= 0 Then
                        MsgBox "Unable to Delete MC Registration Info!!!", vbCritical, "Warning"
                        GoTo endProc
                     End If
                     
                     ' after undoing registration transaction, process its expenses
                     lsSQL = "DELETE FROM MC_Registration_Expense" & _
                              " WHERE sReferNox = " & strParm(lorsx("sTransNox"))
                     p_oAppDrivr.RegisDelete lsSQL
                     p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
                     
                     If Row <= 0 Then
                        MsgBox "Unable to Delete MC Registration Expense Info!!!", vbCritical, "Warning"
                        GoTo endProc
                     End If
                     
                     ' lastly, undo insurance transaction
                     lsSQL = "DELETE FROM MC_Insurance" & _
                              " WHERE sSerialID = " & strParm(lors("sSerialID")) & _
                                 " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                                 " AND sSourceCD = " & strParm("Sl")
                     
                     p_oAppDrivr.RegisDelete lsSQL
                     p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
                     
                     If Row <= 0 Then
                        MsgBox "Unable to Delete Customer Insurance Info!!!", vbCritical, "Warning"
                        GoTo endProc
                     End If
                     
                     If lors("cMotorNew") = xeNo Then
                        ' if a deed of sale expense is created then remove it
                        lsSQL = "DELETE FROM MC_DeedSale_Expense" & _
                                 " WHERE sSerialID = " & strParm(lors("sSerialID")) & _
                                   " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                                   " AND sSourceCD = " & strParm("SDS")
                                    
                        p_oAppDrivr.RegisDelete lsSQL
                        p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
                     End If
                  ElseIf lors("nInsAmtxx") > 0 And lors("cMotorNew") = xeNo Then
                     lsSQL = "DELETE FROM MC_DeedSale_Expense" & _
                              " WHERE sSerialID = " & strParm(lors("sSerialID")) & _
                                " AND sReferNox = " & strParm(.Master("sTransNox")) & _
                                " AND sSourceCD = " & strParm("SDS")
                     
                     p_oAppDrivr.RegisDelete lsSQL
                     p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
                  End If
               End If
               
               If UpdateImpounded(lors("sSerialID"), True) = False Then GoTo endProc
            End If
            lors.MoveNext
         Loop Until lors.EOF
               
         unSaveMCUnits = loMCTrans.Sales(.Master("sTransNox"), True)
         If unSaveMCUnits = False Then GoTo endProc
      Next
   End With

   unSaveMCUnits = True
   
endProc:
   Set loMCTrans = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & "lors" & " )"
   GoTo endProc
End Function

Private Function undoGiveAways() As Boolean
   Dim loSpareparts As clsSpareparts
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim Row As Long
   
   lsOldProc = "UndoGiveAway"
   'On Error Goto errProc
   undoGiveAways = False
   
   Set lors = New Recordset
   Set loSpareparts = New clsSpareparts
   Set loSpareparts.AppDriver = p_oAppDrivr
   
   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  sPartsIDx" & _
                  ", nGivenxxx" & _
               " FROM MC_SO_GiveAways" & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then
         undoGiveAways = True
         GoTo endProc
      End If
      
      loSpareparts.InitTransaction
      loSpareparts.Branch = p_sBranchCd
      loSpareparts.TransactDate = .MasterOriginal("dTransact")
      
      pnCtr = 1
      Do
         loSpareparts.PartsID = lors("sPartsIDx")
         loSpareparts.Quantity = lors("nGivenxxx")
         If loSpareparts.SalesGiveAway(.Master("sTransNox"), True) = False Then GoTo endProc
   
         lors.MoveNext
      Loop Until lors.EOF
      
      lsSQL = "DELETE FROM MC_SO_GiveAways" & _
               " WHERE sTransNox = " & strParm(.Master("sTransNox"))
      
      p_oAppDrivr.RegisDelete lsSQL
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row = 0 Then
         MsgBox "Unable to Undo MC Sales GiveAways Transaction!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   undoGiveAways = True
   
endProc:
   Set loSpareparts = Nothing
   Set lors = Nothing

   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function saveApplication() As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim Row As Long
   
   lsOldProc = "SaveApplication"
   'On Error Goto errProc
   saveApplication = False
   
   With p_oMCARAct
      lsSQL = "INSERT INTO MC_Credit_Application " & _
                  "( sTransNox" & _
                  ", cApplType" & _
                  ", sClientID" & _
                  ", dAppliedx" & _
                  ", sCredCoun" & _
                  ", sAgentIDx" & _
                  ", cUnitAppl" & _
                  ", sUnitAppl" & _
                  ", nGrossPrc" & _
                  ", nPNValuex" & _
                  ", nDownPaym" & _
                  ", sModelIDx" & _
                  ", nMonAmort" & _
                  ", nAcctTerm"
      lsSQL = lsSQL & _
                  ", sQMatchNo" & _
                  ", sNickName" & _
                  ", sLenStayx" & _
                  ", sExAddres" & _
                  ", sExTownID" & _
                  ", sSpouseID" & _
                  ", sSpLenSty" & _
                  ", sCoMakrID" & _
                  ", sRelation" & _
                  ", sRemarksx" & _
                  ", cTranStat" & _
                  ", sApproved" & _
                  ", sModified" & _
                  ", dModified"
      lsSQL = lsSQL & ") VALUES ( " & _
                  strParm(.Fields("sApplicNo")) & _
                  ", " & strParm("0") & _
                  ", " & strParm(p_oBaseClas.Master("sClientID")) & _
                  ", " & dateParm(p_oBaseClas.Master("dTransact")) & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("0") & _
                  ", " & strParm("") & _
                  ", " & .Fields("nGrossPrc") & _
                  ", " & .Fields("nPNValuex") & _
                  ", " & .Fields("nDownPaym") & _
                  ", " & strParm(p_oBaseClas.Detail(1, "sModelIDx")) & _
                  ", " & .Fields("nMonAmort") & _
                  ", " & .Fields("nAcctTerm")
      lsSQL = lsSQL & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("") & _
                  ", " & strParm("«PreCreated : " & p_oBaseClas.Master("sTransNox") & "»") & _
                  ", " & strParm(xeStateUnknown) & _
                  ", " & strParm("") & _
                  ", " & strParm(Encrypt(p_oAppDrivr.UserID)) & _
                  ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
                  
      p_oAppDrivr.Connection.Execute lsSQL, Row, adCmdText
      
      If Row <= 0 Then
         MsgBox "Unable to Save Created Application Entry!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   saveApplication = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function getReplacedUnit(ByVal Row As Integer) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "getReplacedUnit"
   'On Error Goto errProc
   getReplacedUnit = False

   With p_oBaseClas
      lsSQL = "SELECT" & _
                  "  a.sSerialID" & _
                  ", e.dTransact" & _
                  ", a.sEngineNo" & _
                  ", a.sFrameNox" & _
                  ", b.sModelNme" & _
                  ", c.sColorNme" & _
                  ", d.nUnitPrce" & _
                  ", d.nInsAmtxx" & _
                  ", a.sMCInvIDx" & _
               " FROM MC_Serial a" & _
                  ", MC_Model b" & _
                  ", Color c" & _
                  ", MC_SO_Detail d" & _
                  ", MC_SO_Master e" & _
               " WHERE a.sModelIDx = b.sModelIDx" & _
                  " AND a.sColorIDx = c.sColorIDx" & _
                  " AND a.sSerialID = d.sSerialID" & _
                  " AND d.sTransNox = e.sTransNox" & _
                  " AND a.cLocation = " & strParm(xeLocCustomer) & _
                  " AND e.sClientID = " & strParm(.Master("sClientID"))
   End With
      
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   ' check if there's unit bought by the customer
   If lors.EOF Then
      MsgBox "Customer Did Not Bought Motorcycle Here!!!" & _
               "Motorcycle Replacement is Not Allowed!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If lors.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, lors, _
                  "sEngineNo»sModelNme»nUnitPrce»nInsAmtxx", _
                  "Engine No»Model»Unit Price»Insurance")
      
      If lsSQL = Empty Then
         MsgBox "No Selection Was Made!!!" & _
                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      lors.MoveFirst
      lors.Find "sSerialID = " & strParm(Left(lsSQL, 10))
   End If
   
   With p_oBaseClas
      .Detail(Row, "sReplMCID") = lors("sSerialID")
      .Detail(Row, "nReplAmtx") = lors("nUnitPrce")
      .Detail(Row, "nReplIAmt") = lors("nInsAmtxx")
      .Detail(Row, "sRMCInvID") = lors("sMCInvIDx")
   End With
   
   getReplacedUnit = True
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & Row & " )"
   GoTo endProc
End Function

Private Sub Class_Initialize()
   Set p_oMCARAct = New Recordset
   p_cTranStat = xeStateUnknown
End Sub

Private Sub Class_Terminate()
   Set p_oMCARAct = Nothing
   Set p_oBaseClas = Nothing
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsOldProc As String
   
   lsOldProc = "p_oBaseClas_InitMaster"
   'On Error Goto errProc

   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = Date
      .Master("sDRNoxxxx") = GetNextReference(.MasterTable, "sDRNoxxxx", "dTransact", _
                                 "LEFT(sTransNox, 2)", p_sBranchCd, p_oAppDrivr.Connection)
      .Master("nTranTotl") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("sORNoxxxx") = ""
      .Master("sApproved") = ""
      .Master("sRemarksx") = ""
      .Master("cPaymForm") = 0
      .Master("cTranType") = 1
      .Master("cCustType") = 1
      .Master("cTranStat") = xeStateOpen
      .Master("sApplicNo") = ""
      .Master("sClientID") = ""
      .Master("sSalesInv") = ""
      
      .Detail(0, "sTransNox") = .Master("sTransNox")
      .Detail(0, "nEntryNox") = 0
      .Detail(0, "sEngineNo") = ""
      .Detail(0, "sFrameNox") = ""
      .Detail(0, "sModelNme") = ""
      .Detail(0, "sColorNme") = ""
      .Detail(0, "sInsTypNm") = ""
      .Detail(0, "nInsAmtxx") = 0#
      .Detail(0, "nUnitPrce") = 0#
      .Detail(0, "sWarrntNo") = ""
      .Detail(0, "sRemarksx") = ""
      .Detail(0, "cMotorNew") = ""
      .Detail(0, "sMCInvIDx") = ""
      .Detail(0, "sSerialID") = ""
      .Detail(0, "sInsTypID") = ""
      .Detail(0, "cMotorUnt") = 0
      .Detail(0, "sReplMCID") = ""
      .Detail(0, "nReplAmtx") = 0#
      .Detail(0, "nReplIAmt") = 0#
      .Detail(0, "sRMCInvID") = ""
   End With
   
   With p_axReceipt
      .TransNox = GetNextCode("Receipt_Master", "sTransNox", True, _
                     p_oAppDrivr.Connection, True, p_sBranchCd)
      .ORNoxxxx = GetNextReference("Receipt_Master", "sORNoxxxx", "dTransact", _
                                 "LEFT(sTransNox, 2)", p_sBranchCd, p_oAppDrivr.Connection)
      .TranTotl = 0#
      .CashAmtx = 0#
      .CheckAmt = 0#
      .Remarksx = ""
   End With
   
   For pnCtr = 0 To UBound(p_axChecks)
      p_axChecks(pnCtr).AcctNoxx = ""
      p_axChecks(pnCtr).BankName = ""
      p_axChecks(pnCtr).CheckNox = ""
      p_axChecks(pnCtr).CheckDte = Date
      p_axChecks(pnCtr).Amountxx = 0
   Next

   Set p_oMCARAct = New Recordset
   p_oMCARAct.Open p_sActQuery & " AND 0 = 1", _
                     p_oAppDrivr.Connection, , adLockOptimistic, adCmdText
   
   p_sFullName = ""
   p_sAddressx = ""
   p_bGiveAway = True
   pbCreatApp = False

endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & " )"
End Sub

Private Sub p_oBaseClas_LoadOthers()
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "p_oBaseClas_LoadOthers"
   'On Error Goto errProc
   
   pbChecksOK = True
   pbRegisOK = True
   With p_oBaseClas
      p_sFullName = .Master("xFullName")
      p_sAddressx = .Master("xAddressx")
      
      Set lors = New Recordset
      
      If .Master("cTranStat") = xeStateCancelled Then GoTo endProc
      If .Master("nAmtPaidx") > 0 Then
         lsSQL = "SELECT" & _
                     "  sTransNox" & _
                     ", nCashAmtx" & _
                     ", nCheckAmt" & _
                     ", sRemarksx" & _
                  " FROM Receipt_Master" & _
                  " WHERE sReferNox = " & strParm(.Master("sTransNox"))
         lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         If lors.EOF Then
            Err.Raise 32006, "clsMCSales", "Unable to retrieve receipt master info"
            Err.Clear
         End If
         
         p_axReceipt.CashAmtx = lors("nCashAmtx")
         p_axReceipt.CheckAmt = lors("nCheckAmt")
         p_axReceipt.ORNoxxxx = .Master("sORNoxxxx")
         p_axReceipt.Remarksx = lors("sRemarksx")
         p_axReceipt.TransNox = lors("sTransNox")
         p_axReceipt.TranTotl = .Master("nTranTotl")
         
         If p_axReceipt.CheckAmt > 0 Then
            lsSQL = "SELECT" & _
                        "  sCheckNox" & _
                        ", sAcctNoxx" & _
                        ", sBankName" & _
                        ", dCheckDte" & _
                        ", nAmountxx" & _
                        ", cDepositd" & _
                     " FROM Checks_Received" & _
                     " WHERE sReferNox = " & strParm(lors("sTransNox")) & _
                     " ORDER BY sTransNox"
            lors.Close
            lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
 
            If lors.EOF Then
               Err.Raise 32006, "clsMCSales", "Unable to retrieve check detail info"
               Err.Clear
            End If
            
            pnCtr = 0
            Do
               p_axChecks(pnCtr).CheckNox = lors("sCheckNox")
               p_axChecks(pnCtr).AcctNoxx = lors("sAcctNoxx")
               p_axChecks(pnCtr).BankName = lors("sBankName")
               p_axChecks(pnCtr).CheckDte = lors("dCheckDte")
               p_axChecks(pnCtr).Amountxx = lors("nAmountxx")
               
               If lors("cDepositd") = xeYes Then pbChecksOK = False
               pnCtr = pnCtr + 1
               lors.MoveNext
            Loop Until lors.EOF
         End If
      End If
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sInsTypID") <> Empty Then
            lsSQL = "SELECT sPolicyNo FROM MC_Insurance" & _
                     " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
                       " AND sSourceCD = " & "'Sl'" & _
                       " AND sSerialID = " & strParm(.Detail(pnCtr, "sSerialID"))
            
            If lors.State = adStateOpen Then lors.Close
                     
            lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
            If Not lors.EOF Then
               If Not IsNull(lors("sPolicyNo")) Then
                  If lors("sPolicyNo") <> Empty Then pbRegisOK = False
               End If
            End If
         End If
      Next
      
      Set p_oMCARAct = New Recordset
      If .Master("sApplicNo") <> Empty Then
         If .Master("sApplicNo") <> "N-O-N-E" Then
            lsSQL = "SELECT" & _
                        "  a.sAcctNmbr" & _
                        ", a.sApplicNo" & _
                        ", c.sRouteNme" & _
                        ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) AS xCollectr" & _
                        ", CONCAT(e.sLastName, ', ', e.sFrstName, ' ', e.sMiddName) AS xManagerx" & _
                        ", f.sBranchNm xSBranchx" & _
                        ", a.dFirstPay" & _
                        ", a.nAcctTerm" & _
                        ", a.dDueDatex" & _
                        ", a.nGrossPrc" & _
                        ", a.nDownPaym" & _
                        ", a.nCashBalx" & _
                        ", a.nPNValuex" & _
                        ", a.nMonAmort" & _
                        ", a.nPenaltyx" & _
                        ", a.nRebatesx" & _
                        ", c.sCollctID" & _
                        ", c.sManagrID" & _
                        ", g.sRouteIDx" & _
                        ", a.nDownTotl" & _
                        ", a.nABalance"

            lsSQL = lsSQL & _
                        ", a.cAcctStat cTranStat" & _
                        ", f.sBranchCd" & _
                        ", b.sTownIDxx" & _
                        ", a.sRemarksx" & _
                        ", a.cLoanType" & _
                     " FROM MC_AR_Master a" & _
                        ", Client_Master b" & _
                        ", Route_Area c" & _
                        ", Employee_Master d" & _
                        ", Employee_Master e" & _
                        ", Branch f" & _
                        ", Route_Area_Town g" & _
                        ", TownCity i" & _
                        ", Province j" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND b.sTownIDxx = g.sTownIDxx" & _
                        " AND g.sRouteIDx = c.sRouteIDx" & _
                        " AND c.sCollctID = d.sEmployID" & _
                        " AND c.sManagrID = e.sEmployID" & _
                        " AND c.sBranchCd = f.sBranchCd" & _
                        " AND b.sTownIDxx = i.sTownIDxx" & _
                        " AND i.sProvIDxx = j.sProvIDxx" & _
                        " AND a.sApplicNo = " & strParm(.Master("sApplicNo"))
         Else
            lsSQL = GetAccountQuery
         End If
      Else
         lsSQL = p_sActQuery & " AND 0 = 1"
      End If

      p_oMCARAct.Open lsSQL, p_oAppDrivr.Connection, , adLockOptimistic, adCmdText
   End With
   
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveComplete()
   Dim lsOldProc As String
   
   lsOldProc = "p_oBaseClas_SaveComplete"
   'On Error Goto errProc

   If pbModified = False Then GoTo endProc
   If p_bGiveAway = False Then GoTo endProc
   
   saveGiveAways
   saveAccessories
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "p_oBaseClas_WillCancel"
   'On Error Goto errProc
   
   Cancel = Not UndoTransaction

endProc:
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & " )"
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsOldProc = "p_oBaseClas_WillDelete"
   'On Error Goto errProc
   
   Cancel = Not UndoTransaction
   
endProc:

   Exit Sub
errProc:
    ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim loClient As clsClient
   Dim lsOldProc As String
   Dim lbCancel As Boolean
   
   lsOldProc = "p_oBaseClas_WillSave"
   'On Error Goto errProc
   Cancel = True
   
   pbModified = True
   With p_oBaseClas
      If .isMasterModify = False And .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If
      
      RaiseEvent WillSave(lbCancel)
      If lbCancel Then GoTo endProc
      
      If .EditMode = xeModeUpdate Then
         If UndoTransaction = False Then GoTo endProc
      End If
         
      Set loClient = New clsClient
      Set loClient.AppDriver = p_oAppDrivr
      
      If pbCreatApp = True Then
         If saveApplication = False Then GoTo endProc
      End If
      
      If saveAccount = False Then GoTo endProc
      
      If .ItemCount = 1 Then
         If saveMCUnit = False Then GoTo endProc
      Else
         If saveMCUnits = False Then GoTo endProc
      End If

      ' check if transaction detail exist
      If .Detail(0, "sEngineNo") = "" And .Detail(0, "sRemarksx") = Empty Then
         MsgBox "Invalid Transaction Detail Detected!!!" & vbCrLf & _
                  "Verify your Entry then Try Again!!!" & vbCrLf & _
                  .Detail(0, "sEngineNo"), vbCritical, "Warning"
         GoTo endProc
      End If

      Debug.Print "Hola"
      If .Master("nAmtPaidx") > 0 Then
         .Master("cTranStat") = xeStateClosed
         Cancel = Not SaveReceipt
         If Cancel Then GoTo endProc
      End If

      ' only institutional sales has its transaction ledger
      If .Master("cCustType") = 0 Then
         ' update the client ledger
         loClient.InitTransaction
         loClient.Branch = p_sBranchCd
         loClient.ClientType = xeClientCustomer
         loClient.ClientID = .Master("sClientID")
         loClient.DebitAmount = .Master("nTranTotl")
         loClient.CreditAmount = .Master("nAmtPaidx")
         loClient.SourceSystem = "MC"
         loClient.TransactDate = .Master("dTransact")
         If loClient.Sales(.Master("sTransNox")) = False Then GoTo endProc
      End If
   End With
   
   Cancel = False
   
endProc:
   Set loClient = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
   GoTo endProc
End Sub
