VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCSales"
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
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' Note:
'     Kalyptus [2012.08.15 14:58pm]
'        Replacement
'          + unit price should not be reentered...
'          + insurance amount should be reentered/indicated if customer paid an insurance
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
'  XerSys [ 06/27/2007 01:42 pm ]
'     Separate the Repo Inventory from New Inventory
'  XerSys [ 09/28/2007 02:56 pm ]
'     Include Saving of Giveaways
'  XerSys [ 10/11/2007 03:33 pm ]
'     Include Replacement of Giveaways
'  XerSys [ 10/19/2007 10:06 am ]
'     Adopt the new implementation of Motorcycle Registration Object
'  XerSys [ 01/02/2008 09:35 am ]
'     Implement Advance payment entry.
'  XerSys [ 02/04/2008 01:07 pm ]
'     Incorporate New Loan Receivable class
'  XerSys [ 02/13/2008 10:07 am ]
'     Incorporate the new Loan Application class
'  Kalyptus [ 10/20/2008 02:07 Pm ]
'     Added a feature to select the form of registration to process. Please see getInsurance.
'  Kalyptus [ 02/05/2010 02:07 Pm ]
'     Added the Co-Client feature
'  XerSys [ 06-10-2010 02:17 pm ]
'     Add payment form for Sumisho Sales
'  Kalyptus [ 10-19-2011 11:04 am ]
'     Added saving of client transaction for term
'     Status: For implementation
'  kalyptus [ 08/24/2012 12:02 pm ]
'     Added a capability to reset the status of Provisionary Receipt Master if SI is cancelled
'     or deleted...
'  kalyptus [ 12/18/2012 11:27 am ]
'     Redefined the validation for the amount allocated for Deed-Of-Sale/Registration.
'  XerSys [ 05/08/2013 09:43 am ]
'     Integrate price validation here
'  XerSys [ 08/16/2013 04:48 pm ]
'     Update computation of down payment - only excluding the chattel mortgage fee
'     Allow user to add giveaway to no-giveaway model which includes used motorcycle
'  XerSys [ 08/26/2013 02:23 pm ]
'     Check value of issue value of additional give away
'  XerSys [ 01/27/2014 11:10 am ]
'     Receipt for cash advance payment must be the OR. Adjust retrieving and updating
'        of affected transactions.
'  XerSys [ 03/24/2014 03:57 pm ]
'     Modify the approach in saving advance payment
'  XerSys [ 07/01/2014 08:59 am ]
'     Incorporate verification of Loan Receivable for replacement sales
'  XerSys [ 07/15/2014 02:28 pm ]
'     Allow official receipt to be issued on check payment
'  XerSys [ 05/11/2015 10:50 am ]
'     Incorporate Agent to sales
'  mac[ 10/26/2015 01:30 pm ]
'     Incorporate MAPFRE Insurance to sales
'  mac [ 12/24/2015 01:00 pm ]
'     Connects MC Sales to MC Product Inquiry(saveInquiry; Inquiry Proc)
'  jeff [ 08/24/2019 01:25 pm ]
'    add validation for hot item model
'  mac [ 08/24/2020 01:00 pm ]
'     add approval code for sale commission
'        maximum commission will be 1000 only
'        HIM must not have commision
'  mac [ 05/28/2022 01:00 pm ]
'     add control for maximum SPA issued by AH/TLM
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsMCSales"
Private Const pxeSystemCode As String = "MC"
Private Const pxeChattelM As String = "ChattelM"
Private Const pxeMinReg_DED As Currency = 600
Private Const pxeTransNox As Integer = 12
Private Const pxeAdded = 3
Private Const pxeMCSales = "MCSl"
Private Const pxeJavaPath As String = "d:\GGC_Java_Systems\"

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
Private p_oClient As clsNeoClient
Private p_oSCA_Const As New clsCodeGenConst
Private p_oGift As GiftCheck

Private p_oCoClt1 As clsNeoClient
Private p_oCoClt2 As clsNeoClient
Private p_oGCCltx As clsNeoClient

Private p_oClientInfo As ClientInfo
Private p_oCoMkrInfo1 As ClientInfo
Private p_oCoMkrInfo2 As ClientInfo

Private p_oAccident As clsAccidentInsurance

Private p_oAgent As clsNeoClient
Private p_oAgentInfo As Recordset
Private p_oInsurance As Recordset
Private p_oGCardAppl As Recordset

Private p_bNewAgent As Boolean
Private p_bGCardAgent As Boolean
Private p_sAgentSQL As String
Private p_nCommAmtx As Double
Private p_sReferNox As String
Private p_cReleased As String
Private p_sReleased As String
Private p_dReleased As Date

Private p_sCoCltNm1 As String
Private p_sCoCltNm2 As String
Private p_sAgentNme As String
Private p_sInsurnce As String

Private p_oLRMaster As clsLRMaster
Private p_oLRApplic As ggcLRApplication.clsLRApplication
Private p_oGiveAway As clsGiveAway
Private p_oMCPrice As clsMCPriceList
Private p_oAdvPaym As Recordset
Private p_bValidate As Boolean

Private p_xChecks As CheckInfo
Private p_axReceipt As ReceiptInfo

Private p_sBranchCd As String
Private p_cTranStat As xeTransactionStatus
Private p_sFullName As String
Private p_sAddressx As String
Private p_sApplicNo As String
Private p_sPRNoxxxx As String
Private p_sAdvRefer As String
Private p_nAdvPaymx As Double
Private p_nAdvForOR As Double
Private p_bIsDelete As Boolean
Private p_nChattelM As Double
Private p_bIssuedOR As Boolean
Private pbModified As Boolean
Private pbInitTran As Boolean

Private psModelIDs As String
Private pbChecksOK As Boolean
Private pbHasGAway As Boolean
Private pbHotItemx As Boolean
Private pnCtr As Integer

Private p_sSQLGvwy As String
Private p_sSQLRebt As String
Private p_sSQLSRPx As String
Private p_sSQLFPay As String
Private p_sSQLCSRx As String 'mac 2020.07.03 sql for approval code entry of sales w/o PNP CSR
Private p_sSQgocas As String 'mac 2020.07.31 sql for updating the status of Online Credit Application
Private p_sSQLComm As String 'mac 2020.08.24 sql for sales commision approval code
Private p_sSQLDate As String 'mac 2021.09.21 sql for back/advance date approval code

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get GiveAways() As clsGiveAway
   Set GiveAways = p_oGiveAway
End Property

Property Set GiftCoupon(vValue As GiftCheck)
   Debug.Print "clsMCSales.GiftCoupon.Detail(1, gcertnox) = " & vValue.Detail(1, "gcertnox")
   Set p_oGift = vValue
End Property

Property Get GiftCoupon() As GiftCheck
   Set GiftCoupon = p_oGift
End Property

Property Get ClientInfo() As ClientInfo
   Set ClientInfo = p_oClientInfo
End Property

Property Get CoMakerInfo1() As ClientInfo
   Set CoMakerInfo1 = p_oCoMkrInfo1
End Property

Property Get CoMakerInfo2() As ClientInfo
   Set CoMakerInfo2 = p_oCoMkrInfo2
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_oBaseClas.EditMode
End Property

Property Get ARAccount(ByVal Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   ARAccount = p_oLRMaster.Master(Index)
   
End Property

Property Let ARAccount(ByVal Index As Variant, ByVal Value As Variant)
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   p_oLRMaster.Master(Index) = Value
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get Checks(ByVal Index As Variant) As Variant
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "schecknox"
      Checks = p_xChecks.CheckNox
   Case 1, "sacctnoxx"
      Checks = p_xChecks.AcctNoxx
   Case 2, "sbankidxx"
      Checks = p_xChecks.BankIDxx
   Case 3, "dcheckdte"
      Checks = p_xChecks.CheckDte
   Case 4, "namountxx"
      Checks = p_xChecks.Amountxx
   Case 5, "sprnoxxxx"
      Checks = p_xChecks.PRNoxxxx
   Case Else
      MsgBox Index & " is not a valid check field"
   End Select
End Property

Property Let Checks(ByVal Index As Variant, ByVal Value As Variant)
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, "schecknox"
      p_xChecks.CheckNox = Value
   Case 1, "sacctnoxx"
      p_xChecks.AcctNoxx = Value
   Case 2, "sbankidxx"
      p_xChecks.BankIDxx = Value
   Case 3, "dcheckdte"
      p_xChecks.CheckDte = Value
   Case 4, "namountxx"
      p_xChecks.Amountxx = Value
   Case 5, "sprnoxxxx"
      p_xChecks.PRNoxxxx = Value
   Case 6, "bissuedor"
      p_bIssuedOR = Value
   Case Else
      MsgBox Index & " is not a valid check field"
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
   Case 7, "nunitprce"
      If IsNumeric(Value) = False Then Exit Property
      
      p_oBaseClas.Detail(Row, Index) = CDbl(Value)
      Call computeTotal
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      Master = p_sFullName
   Case 4, "xaddressx"
      Master = p_sAddressx
   Case 26, "xCoCltNm1"
      Master = p_sCoCltNm1
   Case 27, "xCoCltNm2"
      Master = p_sCoCltNm2
   Case 40, "nadvpaymx"
      Master = p_nAdvPaymx
   Case 31, "sagentnme"
      Master = p_sAgentNme
   Case "slastname"
      Master = p_oClient.Master("sLastName")
   Case "sfrstname"
      Master = Trim(p_oClient.Master("sFrstname") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " "))
   Case "smiddname"
      Master = p_oClient.Master("sMiddName")
   Case Else
      Master = p_oBaseClas.Master(Index)
   End Select
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   With p_oBaseClas
      If Not IsNumeric(Index) Then Index = LCase(Index)
      Select Case Index
      ' enable this if the end of day transaction is okie
      Case 1, "dtransact"
         If Not IsDate(Value) Then
            .Master(Index) = p_oAppDrivr.ServerDate
         Else
            .Master(Index) = CDate(Value)
         End If
         
         RaiseEvent MasterRetrieved(Index)
      Case 3, "xfullname"
         Call getCustomer(Value, False)
      Case 4, "xaddressx"
         p_sAddressx = Value
      Case 5, "stermname"
         getTerm Value, False
      Case 6, "namtpaidx"
         .Master(Index) = Value
      Case 10, "sapplicno"
         Select Case .Master("cPaymForm")
         Case "0", "3", "4", "5"   ' Cash; Term; Term with Registration; Finance
            .Master(Index) = ""
         Case Else
            ' search credit application
            Call getApplication
         End Select
      Case 24, "sadvrefer"
         Select Case Value
         Case xeYes, xeNo
            p_nAdvPaymx = 0#
            If Value = xeYes Then
               If getAdvancePayment = False Then Value = xeNo
            End If
            
            .Master(Index) = Value
            RaiseEvent MasterRetrieved(40)
         Case Else
            .Master(Index) = xeNo
         End Select
      Case 26, "xcocltnm1"
         Call getCoClient1(Value, False)
      Case 27, "xcocltnm2"
         Call getCoClient2(Value, False)
      Case 31, "sagentnme"
         Call getAgent(Value, False)
      Case Else
         .Master(Index) = Value
      End Select
   End With
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
End Property

Property Get MasFldSize(ByVal Index As Integer) As Variant
   On Error Resume Next

   If pbInitTran = False Then Exit Property
   
   MasFldSize = p_oBaseClas.MasFldSize(Index)
End Property

Property Get PRNumber() As String
   PRNumber = p_sPRNoxxxx
End Property

Property Let PRNumber(ByVal Value As String)
   p_sPRNoxxxx = Value
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
   Case 6, "nadvforor"
      Receipt = p_nAdvForOR
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
   Case 6, "nadvforor"
   End Select
End Property

Property Let TransStatus(ByVal Value As xeTransactionStatus)
   p_cTranStat = Value
End Property

Function InitTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "InitTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

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
      .DetailTable = "MC_SO_Detail"
      
      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDRNoxxxx" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) AS xFullName" & _
                        ", CONCAT(IF(IFNull(b.sHouseNox, '') = '', '', CONCAT(b.sHouseNox, ' ')), b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) AS xAddressx" & _
                        ", e.sTermName" & _
                        ", a.nAmtPaidx" & _
                        ", a.sORNoxxxx" & _
                        ", a.sApproved" & _
                        ", a.sRemarksx" & _
                        ", a.sApplicNo" & _
                        ", a.sSalesInv" & _
                        ", a.nTranTotl" & _
                        ", a.cPaymForm" & _
                        ", a.cTranType" & _
                        ", a.cCustType" & _
                        ", a.cTranStat" & _
                        ", a.sClientID" & _
                        ", a.sModified" & _
                        ", a.dModified"
                        
      .MasterQuery = .MasterQuery & _
                        ", b.sCompnyNm" & _
                        ", c.sTownIDxx" & _
                        ", a.sTermIDxx" & _
                        ", a.dDueDatex" & _
                        ", a.sAdvRefer" & _
                        ", a.nEntryNox" & _
                        ", CONCAT(f.sLastName, ', ', f.sFrstName, IF(IFNull(f.sSuffixNm, '') = '', ' ', CONCAT(' ', f.sSuffixNm, ' ')), f.sMiddName) AS xCoCltNm1" & _
                        ", CONCAT(g.sLastName, ', ', g.sFrstName, IF(IFNull(g.sSuffixNm, '') = '', ' ', CONCAT(' ', g.sSuffixNm, ' ')), g.sMiddName) AS xCoCltNm2" & _
                        ", a.sCoCltID1" & _
                        ", a.sCoCltID2" & _
                        ", a.cChngOwnr" & _
                        ", h.sClientID xAssurdID" & _
                        ", CONCAT(i.sLastName, ', ', i.sFrstName, IF(IFNull(i.sSuffixNm, '') = '', ' ', CONCAT(' ', i.sSuffixNm, ' ')), i.sMiddName) AS xAssurdNm" & _
                        ", h.sTransNox xInsurTrn" & _
                        ", a.cPromoTyp" & _
                        ", f.sBrgyIDxx"
                        
      .MasterQuery = .MasterQuery & _
                     " FROM " & .MasterTable & " a" & _
                           " LEFT JOIN Term e" & _
                              " ON a.sTermIDxx = e.sTermIDxx" & _
                           " LEFT JOIN Client_Master f" & _
                              " ON a.sCoCltID1 = f.sClientID" & _
                           " LEFT JOIN Client_Master g" & _
                              " ON a.sCoCltID2 = g.sClientID" & _
                           " LEFT JOIN MC_Accident_Insurance h" & _
                              " ON a.sTransNox = h.sSourceNo" & _
                                 " AND h.sSourceCd = 'MCSl'" & _
                           " LEFT JOIN Client_Master i" & _
                              " ON h.sClientID = i.sClientID" & _
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
                        ", e.cRegisTyp" & _
                        ", b.nLedgerNo" & _
                        ", b.nLedgerNo xLedgerNo"
                        
      .DetailQuery = .DetailQuery & _
                        ", g.sBrandNme" & _
                        ", c.sModelIDx" & _
                        ", d.sColorIDx" & _
                        ", c.cHotItemx" & _
                        ", b.cCSRValid" & _
                     " FROM " & .DetailTable & " a" & _
                           " LEFT JOIN Insurance_Type e" & _
                              " ON a.sInsTypID = e.sInsTypID" & _
                           " LEFT JOIN MC_Serial b" & _
                                 " JOIN MC_Model c" & _
                                    " ON b.sModelIDx = c.sModelIDx" & _
                           " LEFT JOIN Brand g" & _
                                 " ON c.sBrandIDx = g.sBrandIdx" & _
                                 " JOIN Color d" & _
                                    " ON b.sColorIDx = d.sColorIDx" & _
                              " ON a.sSerialID = b.sSerialID" & _
                           " LEFT JOIN MC_Serial f" & _
                              " ON a.sReplMCID = f.sSerialID" & _
                     " ORDER BY a.nEntryNox"
   
      .Reference = "a.sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction
      
      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sDRNoxxxx" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
                           
      'iMac 2019-02-11
      '  allow telemarketing to search mc sales with no branch validation
      If LCase(p_oAppDrivr.ProductID) = "lrtrackr" And _
         p_oAppDrivr.ClientID = "GGC_BM0T1" Then
         
         .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", a.sDRNoxxxx" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, IF(IFNull(b.sSuffixNm, '') = '', ' ', CONCAT(' ', b.sSuffixNm, ' ')), b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                           IIf(p_cTranStat <> xeStateUnknown, " AND a.cTranStat = " & p_cTranStat, "") & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      End If
      
      .BrowseColumn(0) = "sDRNoxxxx»xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sDRNoxxxx»CONCAT(b.sLastName, ', ', b.sFrstName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "DR No»Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@»@»@»MMMM DD, YYYY"
      
      ' Query statement for the term field
      .BrowseQuery(22) = "SELECT" & _
                        "  sTermIDxx" & _
                        ", sTermName" & _
                        ", nTermDays" & _
                        ", nDiscDays" & _
                        ", nDiscount" & _
                     " FROM Term" & _
                     " WHERE cRecdStat = " & xeRecStateActive
      .BrowseColumn(22) = "sTermIDxx»sTermName»nTermDays»nDiscDays»nDiscount"
      .BrowseTitle(22) = "Code»Term Name»Term»Disc Days»Disc"
      .BrowseFormat(22) = "@»@»0 Day/s»0 Day/s»0.00"
      
      .BrowseQuery(17) = "SELECT" & _
                           "  a.sClientID" & _
                           ", CONCAT(a.sLastName, ', ', a.sFrstName, IF(IFNull(a.sSuffixNm, '') = '', ' ', CONCAT(' ', a.sSuffixNm, ' ')), a.sMiddName) AS xFullName" & _
                           ", CONCAT(IF(IFNull(a.sHouseNox, '') = '', '', CONCAT(a.sHouseNox, ' ')), a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) AS xAddressx" & _
                        " FROM Client_Master a" & _
                           ", TownCity b" & _
                           ", Province c" & _
                        " WHERE a.sTownIDxx = b.sTownIDxx" & _
                           " AND b.sProvIDxx = c.sProvIDxx" & _
                           " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseColumn(17) = "sClientID»xFullName»xAddressx"
      .BrowseTitle(17) = "Client ID»Customer Name»Address"
      .BrowseCriteria(17) = "a.sClientID»CONCAT(a.sLastName, ', ', a.sFrstName, IF(IFNull(a.sSuffixNm, '') = '', ' ', CONCAT(' ', a.sSuffixNm, ' ')), a.sMiddName)"
         
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sSerialID" & _
                                 ", a.sEngineNo" & _
                                 ", a.sFrameNox" & _
                                 ", b.sModelNme" & _
                                 ", c.sColorNme" & _
                                 ", d.nSelPrice" & _
                                 ", a.sMCInvIDx" & _
                                 ", a.cSoldStat" & _
                                 ", a.nLedgerNo" & _
                                 ", b.sModelIDx" & _
                                 ", c.sColorIDx" & _
                                 ", b.cHotItemx" & _
                                 ", a.cCSRValid" & _
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

'she 2020-07-01
'meron ng configuration po sa get detail at sa xxxotherconfig
'' Note:
''     jef [2020.06.25 08:44am]
''        add PNP & CSR for serial as per sir Marlon
'      .BrowseDetailQuery(0) = .BrowseDetailQuery(0) & _
'                                 " AND a.cPNPClear = " & strParm(xeYes) & _
'                                 " AND a.cCSRValid = " & strParm(xeYes)

                           
      .BrowseDetailColumn(0) = "sEngineNo»sFrameNox»sModelNme»sColorNme"
      .BrowseDetailTitle(0) = "Engine No»Frame No»Model»Color"
      .BrowseDetailCriteria(0) = "a.sEngineNo»a.sFrameNox»b.sModelNme»c.sColorNme"
   End With
   
   ' initialize the client object that will handle the client info
   Set p_oClient = New clsNeoClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
'         .ShowEntry = True
'         .DisplayWarning = True
   End With
   
   Set p_oCoClt1 = New clsNeoClient
   With p_oCoClt1
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
'         .ShowEntry = True
'         .DisplayWarning = True
   End With
   
   Set p_oCoClt2 = New clsNeoClient
   With p_oCoClt2
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
'         .ShowEntry = True
'         .DisplayWarning = True
   End With
   
   Set p_oGCCltx = New clsNeoClient
   With p_oGCCltx
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oLRMaster = New clsLRMaster
   With p_oLRMaster
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .Parent = pxeMODULENAME
      .DisplayConfirmation = False
      If .InitAccount() = False Then GoTo endProc
   End With
   
   Set p_oMCPrice = New clsMCPriceList
   With p_oMCPrice
      Set .AppDriver = p_oAppDrivr
      If Not .InitTransaction Then
         MsgBox "Unable to Initialize Price List Info!", vbCritical, "Warning"
         GoTo endProc
      End If
      p_bValidate = False
   End With
   
   Set p_oLRApplic = New ggcLRApplication.clsLRApplication
   With p_oLRApplic
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .Parent = pxeMODULENAME
      If .InitTransaction() = False Then GoTo endProc
   End With
   
   Set p_oGiveAway = New clsGiveAway
   With p_oGiveAway
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .Parent = pxeMODULENAME
      .DisplayConfirmation = False
      If .InitTransaction() = False Then GoTo endProc
   End With
   
   p_sAgentSQL = "SELECT sClientID" & _
                     ", sBranchCD" & _
                     ", nNoofDepn" & _
                     ", nNoofChld" & _
                     ", sEmployID" & _
                     ", sRelation" & _
                     ", nNoOfCust" & _
                     ", nCommissn" & _
                     ", dCltSince" & _
                     ", sOthersxx" & _
                     ", cRecdStat" & _
                     ", dModified" & _
                  " FROM MC_Agent"
   
   p_nChattelM = CLng(p_oAppDrivr.getConfiguration(pxeChattelM))
   
   Set p_oGCardAppl = Nothing
   
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then NewTransaction = p_oBaseClas.NewTransaction
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "SaveTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If Not pbInitTran Then GoTo endProc
   
   If getAccount() = False Then GoTo endProc
   If getGiveaway() = False Then GoTo endProc
   
   SaveTransaction = p_oBaseClas.SaveTransaction

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "OpenTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   If pbInitTran Then OpenTransaction = p_oBaseClas.OpenTransaction(sTransNo)
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

         'Enable cancel of transaction if Modification date is the same with the current date...
         'kalyptus - 2011-08-29
   If (DateDiff("h", p_oBaseClas.Master("dModified"), p_oAppDrivr.ServerDate) > 24) Then
     If Not p_oAppDrivr.UserLevel = xeEngineer Then
        MsgBox "Unable to delete " & p_oBaseClas.Master("sTransNox") & vbCrLf & _
                      "Transaction was created from the other days other than today.."
         GoTo endProc
      End If
   End If

   p_bIsDelete = True

   If pbInitTran Then DeleteTransaction = p_oBaseClas.DeleteTransaction

endProc:

   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsProcName As String
   Dim lasMaster() As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsProcName = "SearchTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
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
            lsField = "b.sLastName»b.sFrstName"
            lsDescript = "b.sLastName»b.sFrstName"
         End If
      End If
   End If
   
   SearchTransaction = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function PostTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "PostTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then PostTransaction = p_oBaseClas.PostTransaction(sTransNo)

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "CancelTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If pbInitTran Then
      With p_oBaseClas
         If pbChecksOK = False Then
            MsgBox "Check payment of this transaction has been deposited!!!" & _
                     vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
            Exit Function
         End If

         'Enable cancel of transaction if Modification date is the same with the current date...
         'kalyptus - 2011-08-29
         If (DateDiff("h", .Master("dModified"), p_oAppDrivr.ServerDate) > 24) Then
            If Not p_oAppDrivr.UserLevel = xeEngineer Then
               MsgBox "Unable to cancel " & .Master("sTransNox") & vbCrLf & _
                            "Transaction was created from the other days other than today.."
               GoTo endProc
            End If
         End If

         p_bIsDelete = False
         CancelTransaction = .CancelTransaction()
      End With
   End If

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function CloseTransaction(ByVal sTransNo As String) As Boolean
   Dim lsProcName As String
   
   lsProcName = "CloseTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
      
   If pbInitTran Then CloseTransaction = p_oBaseClas.CloseTransaction(sTransNo)

endProc:

   Exit Function
errProc:
    ShowError lsProcName & "( " & sTransNo & " )"
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
      .Detail(pnCtr, "sModelIDx") = ""
      .Detail(pnCtr, "cHotItemx") = 0
   End With
End Function

Function DeleteDetail(ByVal Index As Long) As Boolean
   DeleteDetail = p_oBaseClas.DeleteDetail(Index)
   Call computeTotal
End Function

Function SearchDetail(ByVal Row As Long, _
                        ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
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
   ShowError lsProcName & "( " & Row _
                        & ", " & Index _
                        & ", " & Value & " )"
   GoTo endProc
End Function

Function SearchMaster(ByVal Index As Variant, _
                        Optional Value As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchMaster"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
      
   If pbInitTran = False Then Exit Function
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 3, "xfullname"
      SearchMaster = getCustomer(Value, True)
   Case 5, "stermname"
      SearchMaster = getTerm(Value, True)
   Case 10, "sapplicno"
      Select Case p_oBaseClas.Master("cPaymForm")
      Case "1", "2"
         SearchMaster = getApplication()
      End Select
   Case 26, "xcocltnm1"
      SearchMaster = getCoClient1(Value, True)
   Case 27, "xcocltnm2"
      SearchMaster = getCoClient2(Value, True)
   End Select

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & Index _
                        & ", " & Value & " )"
End Function

Function UpdateTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "UpdateTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   MsgBox "Update is no longer allowed!" & vbCrLf & _
               " If you need further info, please contact Guanzon SEG/SSG..." _
               , vbInformation, "Update Info"
   Exit Function

   With p_oBaseClas

      If OpenTransaction(.Master("sTransNox")) = False Then Exit Function
      
      If pbChecksOK = False Then
         MsgBox "Check payment of this transaction has been deposited!!!" & _
                  vbCrLf & "Transaction is Not Allowed to be Modify!!!", vbCritical, "Warning"
         Exit Function
      End If

'            'Enable update of transaction if Modification date is the same with the current date...
'            'kalyptus - 2011-08-31
'101         If Not (DateDiff("d", .Master("dModified"), p_oAppDrivr.ServerDate) = 0) Then
'103            MsgBox "Unable to update " & .Master("sTransNox") & vbCrLf & _
'                      "Transaction was created from the other days other than today.."
'104            GoTo endProc
'106         End If

      .EditMode = xeModeUpdate
   End With
   
   With p_oGiveAway
      .UpdateTransaction
   End With
   UpdateTransaction = True

endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Sub ViewModify()
   p_oBaseClas.ViewUserModify
End Sub

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Function getGiveaway() As Boolean
   Dim loFormGiveaway As frmGiveaway
   Dim lsProcName As String
   Dim lsModelIDs As String
   Dim lbGetApproval As Boolean
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
   Dim lsCodeType As String
   Dim lsSQL As String
   Dim lnMaxCommision As Double
         
   lsProcName = "getGiveaway"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   p_sSQLGvwy = ""
   lsModelIDs = ""
      
   pbHotItemx = False
   For pnCtr = 0 To p_oBaseClas.ItemCount() - 1
      If p_oBaseClas.Detail(pnCtr, "cMotorNew") = xeYes Then
         lsModelIDs = "»" & p_oBaseClas.Detail(pnCtr, "sModelIDx")
         
         If Not pbHotItemx Then
            If p_oBaseClas.Detail(pnCtr, "cHotItemx") = xeYes Then pbHotItemx = True
         End If
      End If
   Next
   
   lsModelIDs = Mid(lsModelIDs, 2)
   
   If psModelIDs <> lsModelIDs Then
      With p_oGiveAway
         Call .InitGiveaway
         .Master("sTransNox") = p_oBaseClas.Master("sTransNox")
         .Master("dTransact") = p_oBaseClas.Master("dTransact")
         .Master("sClientID") = p_oBaseClas.Master("sClientID")
         .Master("xFullname") = p_sFullName
         .Master("xAddressx") = p_sAddressx
         .Master("cPaymForm") = p_oBaseClas.Master("cPaymForm")
         
         For pnCtr = 0 To p_oBaseClas.ItemCount() - 1
            If p_oBaseClas.Detail(pnCtr, "cMotorNew") = xeYes Then
               Call .AddModel(p_oBaseClas.Detail(pnCtr, "sModelIDx"))
            End If
         Next
         Call .LoadDetail
      End With
   End If
   
   Set loFormGiveaway = New frmGiveaway
   With loFormGiveaway
      Set .AppDriver = p_oAppDrivr
      Set .GiveAway = p_oGiveAway
      .Show 1
      
      If .Cancelled Then GoTo endProc
      
      'kalyptus - 2015.12.23 01:23pm
      'Check for additional giveaway
      With p_oGiveAway
         For pnCtr = 0 To .ItemCount - 1
            If .Detail(pnCtr, "cGAwyStat") = pxeAdded Then
               lbGetApproval = True
               Exit For
            End If
         Next
      End With
                  
      If lbGetApproval Then
         If p_oAppDrivr.Config("cAllowGwy") <> "1" Then lbGetApproval = False
      End If
      
      If lbGetApproval Then
               
         MsgBox "Additional giveway was detected!" & vbCrLf & _
                "This requires an approval code..." & vbCrLf & vbCrLf & _
                "For inqueries, you can contact Guanzon Group's MIS/SSG.", vbOKOnly, "Validation"
               
         'kalyptus - 2014.04.12 03:51pm
         'if below last price then asked for approval code
         If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
            MsgBox "User is not authorized to Approve such giveaway!" & vbCrLf & _
                  "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
            GoTo endProc
         End If
         
         'kalyptus - 2015.12.24 04:13pm
         'TODO: Check the value of the new field in MC_SO_Master table.
         '      The new field is yet to be created!!!
         Select Case p_oBaseClas.Master("cPromoTyp")
         Case "1"
            lsCodeType = p_oSCA_Const.pxeTeleMktg
'         Case "2"
'            lsCodeType = p_oSCA_Const.pxeAdditional
'         Case "3"
'            lsCodeType = p_oSCA_Const.pxeBiyahingFiesta
         End Select
         
         If isValidApproveCode( _
            lsCodeType, _
            "", _
            Mid(lsApprovedCD, 4, 1), _
            Master("dTransact"), _
            p_sFullName, _
            lsApprovedCD) Then
            
            p_oBaseClas.Master("sApproved") = lsApproveName

            lsSQL = "xTransNox"
            p_sSQLGvwy = "INSERT INTO xxxSCA_Usage" & _
                        " SET sTransNox = " & strParm(lsSQL) & _
                           ", sApprCode = " & strParm(lsApprovedCD) & _
                           ", sApproved = " & strParm(lsApproveID) & _
                           ", sSystemCD = " & strParm(lsCodeType) & _
                           ", sSourceNo = " & strParm(Master("sTransNox")) & _
                           ", sSourceCD = " & strParm(pxeMCSales) & _
                           ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
         ElseIf isValidApproveCode( _
            lsCodeType, _
            "", _
            Mid(lsApprovedCD, 4, 1), _
            Master("dTransact"), _
            p_sAgentNme, _
            lsApprovedCD) Then
            
            p_oBaseClas.Master("sApproved") = lsApproveName
            
            
            lsSQL = "xTransNox"
            p_sSQLGvwy = "INSERT INTO xxxSCA_Usage" & _
                        " SET sTransNox = " & strParm(lsSQL) & _
                           ", sApprCode = " & strParm(lsApprovedCD) & _
                           ", sApproved = " & strParm(lsApproveID) & _
                           ", sSystemCD = " & strParm(lsCodeType) & _
                           ", sSourceNo = " & strParm(Master("sTransNox")) & _
                           ", sSourceCD = " & strParm(pxeMCSales) & _
                           ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
         Else
            MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
               "Verify entry then try again!", vbCritical, "Warning"
            GoTo endProc
         End If 'If isValidApproveCode(
      End If 'If lbGetApproval Then
      
      psModelIDs = lsModelIDs
   End With
   
   With p_oGiveAway
      If .ReplaceAmount > 0 Then
         If .GetSPReplacement() = False Then GoTo endProc
      End If
   End With
   
   getGiveaway = True
   
endProc:
   Set loFormGiveaway = Nothing
   
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function getAccount() As Boolean
   Dim loNewAccount As frmNewAccount
   Dim lsProcName As String
   Dim lnMonAmort As Double
   Dim lnMaxRebte As Double
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
   Dim lsSQL As String
   Dim lsCodeType As String
   
   lsProcName = "getAccount"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   With p_oBaseClas
      
      p_sSQLRebt = ""
      
      Select Case .Master("cPaymForm")
      Case "0", "3", "4", "5"
         getAccount = True
         GoTo endProc
      End Select
                     
      'If trantype is replacement then reload application
      If .Master("cTranType") = "2" Then
'         If .Master("sApplicNo") = Empty Or .Master("sApplicNo") = "N-O-N-E" Then
            Call getApplication
            Call p_oLRMaster.InitAccount
      
            p_sApplicNo = .Master("sApplicNo")
            If p_sApplicNo = "N-O-N-E" Then
               Call p_oLRMaster.LoadByCondition( _
                        "a.sClientID = " & strParm(.Master("sClientID")) & _
                        " AND a.sSerialID = " & strParm(.Detail(0, "sReplMCID")))
            ElseIf p_sApplicNo <> "" Then
               Call p_oLRMaster.LoadByCondition("sApplicNo = " & strParm(.Master("sApplicNo")))
               Debug.Print p_oLRMaster.Master("sAcctNmbr")
            End If
'         End If
      End If
      
      If (p_sApplicNo = Empty And p_oLRMaster.Master("sClientID") = "") Or .Detail(0, "cMotorNew") <> xeYes Then
         If p_oLRMaster.NewAccount = False Then GoTo endProc
         
         'now initialize the account
         p_oLRMaster.Master("sClientID") = .Master("sClientID")
         p_oLRMaster.Master("sTownIDxx") = .Master("sTownIDxx")
         p_oLRMaster.Master("sBrgyIDxx") = .Master("sBrgyIDxx")
         p_oLRMaster.Master("sApplicNo") = .Master("sApplicNo")
         p_oLRMaster.Master("sSerialID") = .Detail(0, "sSerialID")
         
         ' XerSys - 2017-08-24
         '  Assign the cMotorNew to LR Master
         p_oLRMaster.Master("cMotorNew") = .Detail(0, "cMotorNew")
         p_oLRMaster.Master("dPurchase") = .Master("dTransact")
         p_oLRMaster.Master("sRouteNme") = ""   ' this will trigger the object to search for route
         If .Master("cPaymForm") = "2" Then    'Installment
            ' XerSys - 2017-08-24
            ' check if loan item is motorcycle
            If p_oLRMaster.Master("sSerialID") = "" Then
               ' Loan is not a motorcycle, do not deduct chattel mortgage fee
               p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx)
            Else
               ' XerSys 2014-05-14
               '  Check if installment is a repo unit
               If p_oBaseClas.Detail(0, "cMotorNew") = xeYes Then
                  If isBNOS(p_oBaseClas.Detail(0, "sSerialID")) Then
                     p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) '- .Detail(0, "nInsAmtxx")
                  Else
                     p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) - p_nChattelM '.Detail(0, "nInsAmtxx")
                  End If
               Else
                  p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) '- .Detail(0, "nInsAmtxx")
               End If
            End If
            
            ' XerSys - 2015-05-11
            '  Remove the commission of agent from the computation of MA
            p_oLRMaster.Master("nDownPaym") = p_oLRMaster.Master("nDownPaym") - p_nCommAmtx
            p_oLRMaster.Master("nDownTotl") = p_oLRMaster.Master("nDownPaym") - p_nCommAmtx
            
            If p_oLRMaster.Master("nAcctTerm") = 0 Then
'               p_oLRMaster.Master("nPNValuex") = .Master("nTranTotl") - _
'                                                   p_oLRMaster.Master("nDownPaym")
               p_oLRMaster.Master("nAcctTerm") = 12
            End If
            
            ' allow user to encode models that are not yet existing in our database
            If p_bValidate Then
               lnMonAmort = p_oMCPrice.getMonthly(p_oLRMaster.Master("nDownPaym"), p_oLRMaster.Master("nAcctTerm"))
               p_oLRMaster.Master("nPNValuex") = Round(lnMonAmort * p_oLRMaster.Master("nAcctTerm"), 2)
            Else
               p_oLRMaster.Master("nPNValuex") = Round(.Detail(0, "nUnitPrce") * 1.5, 2)
            End If
            p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nPNValuex") + p_oLRMaster.Master("nDownPaym")
            p_oLRMaster.Master("nCashBalx") = 0
         Else
            p_oLRMaster.Master("nPNValuex") = 0
            p_oLRMaster.Master("nAcctTerm") = 0
            p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) ' - .Detail(0, "nInsAmtxx")
            
            ' XerSys - 2015-05-11
            '  Remove the commission of agent from the computation of MA
            p_oLRMaster.Master("nDownPaym") = p_oLRMaster.Master("nDownPaym") - p_nCommAmtx
            p_oLRMaster.Master("nCashBalx") = .Master("nTranTotl") - _
                                                .Detail(0, "nInsAmtxx") - _
                                                p_oLRMaster.Master("nDownPaym")
            p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nCashBalx") + _
                                                p_oLRMaster.Master("nDownPaym")
         End If
      Else
         If .Master("sClientID") <> p_oLRMaster.Master("sClientID") Then
            p_oLRMaster.Master("sClientID") = .Master("sClientID")
         End If
         
         If .Master("sTownIDxx") <> p_oLRMaster.Master("sTownIDxx") Then
            p_oLRMaster.Master("sTownIDxx") = .Master("sTownIDxx")
            p_oLRMaster.Master("sRouteNme") = ""   ' this will trigger the object to search for route
         End If
         
         p_oLRMaster.Master("sApplicNo") = .Master("sApplicNo")
         p_oLRMaster.Master("sSerialID") = .Detail(0, "sSerialID")
         p_oLRMaster.Master("dPurchase") = .Master("dTransact")
         ' XerSys - 2017-08-24
         '  Assign the cMotorNew to LR Master
         p_oLRMaster.Master("cMotorNew") = .Detail(0, "cMotorNew")
         
         ' XerSys - 2014-07-01
         '  Check if sales is a replacement
         If .Master("cTranType") <> "2" Then ' replacement
            If .Master("cPaymForm") = "2" Then ' Installment
               ' XerSys - 2017-08-24
               ' check if loan item is motorcycle
               If p_oLRMaster.Master("sSerialID") = "" Then
                  ' Loan is not a motorcycle, do not deduct chattel mortgage fee
                  p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx)
               Else
                  p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx) _
                                                   - p_nChattelM '.Detail(0, "nInsAmtxx")
               End If
               ' XerSys - 2015-05-11
               '  Remove the commission of agent from the computation of MA
               p_oLRMaster.Master("nDownPaym") = p_oLRMaster.Master("nDownPaym") - p_nCommAmtx

               If p_oLRMaster.Master("nAcctTerm") = 0 Then
                  p_oLRMaster.Master("nAcctTerm") = 12
               End If
               If p_bValidate Then
                  lnMonAmort = p_oMCPrice.getMonthly(p_oLRMaster.Master("nDownPaym"), p_oLRMaster.Master("nAcctTerm"))
               
                  p_oLRMaster.Master("nPNValuex") = Round(lnMonAmort * p_oLRMaster.Master("nAcctTerm"), 2)
               Else
                  p_oLRMaster.Master("nPNValuex") = Round(.Detail(0, "nUnitPrce") * 1.5, 2)
               End If
               p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nPNValuex") + p_oLRMaster.Master("nDownPaym")
               p_oLRMaster.Master("nCashBalx") = 0
            ElseIf .Master("cPaymForm") = "1" And _
                  p_oLRMaster.Master("nCashBalx") = 0 Then ' Cash Balance
               p_oLRMaster.Master("nPNValuex") = 0
               p_oLRMaster.Master("nAcctTerm") = 0
               p_oLRMaster.Master("nDownPaym") = (.Master("nAmtPaidx") + p_nAdvPaymx)
               
               ' XerSys - 2015-05-11
               '  Remove the commission of agent from the computation of MA
               p_oLRMaster.Master("nDownPaym") = p_oLRMaster.Master("nDownPaym") - p_nCommAmtx
'               MsgBox p_oLRMaster.Master("nDownTotl")
               p_oLRMaster.Master("nCashBalx") = .Master("nTranTotl") - _
                                                   p_oLRMaster.Master("nDownPaym")
                                                   '.Detail(0, "nInsAmtxx") -
               p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nCashBalx") + _
                                                   p_oLRMaster.Master("nDownPaym")
            End If
         End If
      End If
   End With
   
   Set loNewAccount = frmNewAccount
   With loNewAccount
      Set .AppDriver = p_oAppDrivr
      Set .LRMaster = p_oLRMaster
      Set .MCPrice = p_oMCPrice
      
      .Commision = p_nCommAmtx   'mac 2020.07.04   added commission to be passed on new account form
      .txtOther(0).Text = p_sFullName
      .txtOther(1).Text = p_sAddressx
      .txtOther(2).Text = IIf(Trim(p_sCoCltNm1) = "", "N-O-N-E", p_sCoCltNm1)
      .txtOther(3).Text = IIf(Trim(p_sCoCltNm2) = "", "N-O-N-E", p_sCoCltNm2)
      'jovan added update of due date once cancelled'
'      .txtField(17).Text = Format(loNewAccount.DueDate, "MMMM DD, YYYY")
      
      .Show 1
      If .Cancelled Then GoTo endProc
   
      'kalyptus - 2015.12.24 04:19pm
      'TODO:Installment sales with rebate of more than 200 should
      '     need an approval code
      'Is it New
      'mac 2022.09.21
      '  ivalidate na rin natin ang rebate kahit repo
      'If p_oBaseClas.Detail(0, "cMotorNew") = xeYes Then
         ' Installment
      If p_oBaseClas.Master("cPaymForm") = "2" Then
         'Is rebate more than 200 pesos
         '2019-03-30
         'jheff
         'creaste field for Maximum Rebate per branch (BranchOthers)
         
         If p_oBaseClas.Detail(0, "cMotorNew") = xeYes Then
            lnMaxRebte = p_oAppDrivr.Config("nMaxRebte")
         Else
            lnMaxRebte = IFNull(p_oAppDrivr.getConfiguration("2HMaxRebte"), 300)
         End If
         
         If p_oLRMaster.Master("nRebatesx") > lnMaxRebte Then
            MsgBox "Rebate given is greater than the supposed rebate." & vbCrLf & _
                     "You will be asked to enter an approval code.", vbInformation, "Notice"
         
            If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
               MsgBox "User is not authorized to Approve such REBATE!" & vbCrLf & _
                     "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
               GoTo endProc
            End If
            
            'kalyptus - 2015.12.24 04:13pm
            'Check the value of the new field in MC_SO_Master table.
            'The new field is yet to be created!!!
            Select Case p_oBaseClas.Master("cPromoTyp")
            Case "2", "3", "0"
               lsCodeType = p_oSCA_Const.pxeAdditional
            Case "1"
               lsCodeType = p_oSCA_Const.pxeTeleMktg
            End Select
            
            If isValidApproveCode( _
               lsCodeType, _
               "", _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               p_sFullName, _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveName

               lsSQL = "xTransNox"
               p_sSQLRebt = "INSERT INTO xxxSCA_Usage" & _
                           " SET sTransNox = " & strParm(lsSQL) & _
                              ", sApprCode = " & strParm(lsApprovedCD) & _
                              ", sApproved = " & strParm(lsApproveID) & _
                              ", sSystemCD = " & strParm(lsCodeType) & _
                              ", sSourceNo = " & strParm(Master("sTransNox")) & _
                              ", sSourceCD = " & strParm(pxeMCSales) & _
                              ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
               
            ElseIf isValidApproveCode( _
               lsCodeType, _
               "", _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               p_sAgentNme, _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveName

               lsSQL = "xTransNox"
               p_sSQLRebt = "INSERT INTO xxxSCA_Usage" & _
                           " SET sTransNox = " & strParm(lsSQL) & _
                              ", sApprCode = " & strParm(lsApprovedCD) & _
                              ", sApproved = " & strParm(lsApproveID) & _
                              ", sSystemCD = " & strParm(lsCodeType) & _
                              ", sSourceNo = " & strParm(Master("sTransNox")) & _
                              ", sSourceCD = " & strParm(pxeMCSales) & _
                              ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If 'If isValidApproveCode(
         End If 'If p_oLRMaster.Master("nRebatesx") > 200
      End If 'If .Master("cPaymForm") = "2"
      'End If 'If p_oBaseClas.Detail(0, "cMotorNew") = xeYes
      
      If p_oBaseClas.Master("cPaymForm") = "2" Then
         'jef
         'validate first pay day
         '2019-08-27
         If IFNull(p_oAppDrivr.getConfiguration("FrstPaym"), "") = xeYes Then
            If IsFirstPayModify(p_oBaseClas.Master("dTransact"), p_oLRMaster.Master("dFirstPay")) Then
               If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
                  MsgBox "User is not authorized to Approve such application!" & vbCrLf & _
                        "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
                  GoTo endProc
               End If
            
               If isValidApproveCode( _
                  p_oSCA_Const.pxeARFirstPayDate, _
                  p_oAppDrivr.BranchCode, _
                  Mid(lsApprovedCD, 4, 1), _
                  p_oLRMaster.Master("dFirstPay"), _
                  Master("sDRNoxxxx"), _
                  lsApprovedCD) Then
                  
                  p_oBaseClas.Master("sApproved") = lsApproveName
                  lsSQL = "xTransNox"
                  p_sSQLFPay = "INSERT INTO xxxSCA_Usage" & _
                              " SET sTransNox = " & strParm(lsSQL) & _
                                 ", sApprCode = " & strParm(lsApprovedCD) & _
                                 ", sApproved = " & strParm(lsApproveID) & _
                                 ", sSystemCD = " & strParm(p_oSCA_Const.pxeARFirstPayDate) & _
                                 ", sSourceNo = " & strParm(Master("sTransNox")) & _
                                 ", sSourceCD = " & strParm(pxeMCSales) & _
                                 ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                                 ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
               Else
                  MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                     "Verify entry then try again!", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         End If
      End If
   End With
   
   'mac 2020.07.31
   '  update the status of Credit Onlin Appication to USED
   If p_sApplicNo <> "" Then
      Dim loRS As Recordset
      
      Set loRS = New Recordset
      
      p_sSQgocas = "SELECT IFNULL(sReferNox, '') sReferNox FROM MC_Credit_Application WHERE sTransNox = " & strParm(p_sApplicNo)
      
      loRS.Open p_sSQgocas, p_oAppDrivr.Connection, , , adCmdText
      Set loRS.ActiveConnection = Nothing
      
      p_sSQgocas = ""
      
      If Not loRS.EOF Then
         If loRS("sReferNox") <> "" Then
            p_sSQgocas = "UPDATE Credit_Online_Application SET" & _
                           "  cTranStat = '2'" & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                        " WHERE sTransNox = " & strParm(loRS("sReferNox"))
         End If
      End If
      
      Set loRS = Nothing
   End If
   'end - mac 2020.07.31
   
   getAccount = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function getCustomer(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getCustomer"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .isLTOID = IFNull(p_oAppDrivr.getConfiguration("LTMS"), "0") = "1"
      If .InitClient() = False Then GoTo endProc
   End With
   
   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(LCase(lsValue)) = Trim(LCase(p_sFullName)) Then GoTo endProc
         lbExist = loClient.SearchClient(lsValue, False)
      Else
         GoTo endWithClear
      End If

      If Not lbExist Then
         lasName = GetSplitedName(lsValue)
         loClient.Master("sLastName") = lasName(0)
         loClient.Master("sFrstName") = lasName(1)
      End If

      If loClient.getClient Then
         Set p_oClient = loClient
      End If
      
      p_oClientInfo.sLastName = p_oClient.Master("sLastName")
      p_oClientInfo.sFrstName = p_oClient.Master("sFrstName")
      p_oClientInfo.sMiddName = p_oClient.Master("sMiddName")
      p_oClientInfo.sSuffixNm = p_oClient.Master("sSuffixNm")
      p_oClientInfo.sBirthDte = p_oClient.Master("dBirthDte")
      p_oClientInfo.sBirthPlc = p_oClient.Master("sBrthTown")
      p_oClientInfo.sHouseNox = p_oClient.Master("sHouseNox")
      p_oClientInfo.sAddresxx = p_oClient.Master("sAddressx")
      p_oClientInfo.sBarangay = p_oClient.Master("sBrgyName")
      p_oClientInfo.sTownxxxx = p_oClient.Master("xTownName")
      p_oClientInfo.sProvince = p_oClient.Master("xProvName")
      p_oClientInfo.sMobileNo = p_oClient.Master("sMobileNo")
      p_oClientInfo.sEmailAdd = p_oClient.Master("sEmailAdd")
      p_oClientInfo.sLTOIDxxx = p_oClient.Master("sLTOIDxxx")
      p_oClientInfo.sFatherNm = p_oClient.Master("sFatherID")
      p_oClientInfo.sMotherNm = p_oClient.Master("sMotherID")
      p_oClientInfo.sOccptnNm = p_oClient.Master("sOccptnNm")
      p_oClientInfo.sCitizenx = p_oClient.Master("sCitizenx")
      p_oClientInfo.sGenderCd = p_oClient.Master("cGenderCd")
      p_oClientInfo.sTaxIDNox = p_oClient.Master("sTaxIDNox")

      .Master("sClientID") = p_oClient.Master("sClientID")
      .Master("sTownIDxx") = p_oClient.Master("sTownIDxx")
      .Master("sBrgyIDxx") = p_oClient.Master("sBrgyIDxx")
      p_oClient.Master("cMCClient") = "1"
      p_sFullName = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
      p_sAddressx = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")

   End With
   
   getCustomer = True
   
endProc:
   RaiseEvent MasterRetrieved(3)
   RaiseEvent MasterRetrieved(4)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sClientID") = ""
   p_oBaseClas.Master("sTownIDxx") = ""
   Set p_oClient = loClient
   p_sFullName = ""
   p_sAddressx = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getCoClient1(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getCoClient1"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
      'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .isLTOID = IFNull(p_oAppDrivr.getConfiguration("LTMS"), "0") = "1"
      If .InitClient() = False Then GoTo endProc
   End With
   
   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(LCase(lsValue)) = Trim(LCase(p_sCoCltNm1)) Then GoTo endProc
         lbExist = loClient.SearchClient(lsValue, False)
      Else
         GoTo endWithClear
      End If

      If Not lbExist Then
         lasName = GetSplitedName(lsValue)
         loClient.Master("sLastName") = lasName(0)
         loClient.Master("sFrstName") = lasName(1)
      End If

      If loClient.getClient Then
         Set p_oCoClt1 = loClient
      End If
      
      p_oCoMkrInfo1.sLastName = p_oCoClt1.Master("sLastName")
      p_oCoMkrInfo1.sFrstName = p_oCoClt1.Master("sFrstName")
      p_oCoMkrInfo1.sMiddName = p_oCoClt1.Master("sMiddName")
      p_oCoMkrInfo1.sSuffixNm = p_oCoClt1.Master("sSuffixNm")
      p_oCoMkrInfo1.sBirthDte = p_oCoClt1.Master("dBirthDte")
      p_oCoMkrInfo1.sBirthPlc = p_oCoClt1.Master("sBrthTown")
      p_oCoMkrInfo1.sHouseNox = p_oCoClt1.Master("sHouseNox")
      p_oCoMkrInfo1.sAddresxx = p_oCoClt1.Master("sAddressx")
      p_oCoMkrInfo1.sBarangay = p_oCoClt1.Master("sBrgyName")
      p_oCoMkrInfo1.sTownxxxx = p_oCoClt1.Master("xTownName")
      p_oCoMkrInfo1.sProvince = p_oCoClt1.Master("xProvName")
      p_oCoMkrInfo1.sMobileNo = p_oCoClt1.Master("sMobileNo")
      p_oCoMkrInfo1.sEmailAdd = p_oCoClt1.Master("sEmailAdd")
      p_oCoMkrInfo1.sLTOIDxxx = p_oCoClt1.Master("sLTOIDxxx")
      p_oCoMkrInfo1.sTaxIDNox = p_oCoClt1.Master("sTaxIDNox")
      p_oCoMkrInfo1.sFatherNm = p_oCoClt1.Master("sFatherID")
      p_oCoMkrInfo1.sMotherNm = p_oCoClt1.Master("sMotherID")
      p_oCoMkrInfo1.sCitizenx = p_oCoClt1.Master(84)
      p_oCoMkrInfo1.sGenderCd = p_oCoClt1.Master(4)
      p_oCoMkrInfo1.sOccptnNm = p_oCoClt1.Master(86)

      .Master("sCoCltID1") = p_oCoClt1.Master("sClientID")
      p_oCoClt1.Master("cMCClient") = "1"
      p_sCoCltNm1 = p_oCoClt1.Master("sLastName") + ", " + p_oCoClt1.Master("sFrstName") + " " + Trim(p_oCoClt1.Master("sSuffixNm")) + IIf(Trim(p_oCoClt1.Master("sSuffixNm")) = "", "", " ") + p_oCoClt1.Master("sMiddName")
   End With
   
   getCoClient1 = True
   
endProc:
   RaiseEvent MasterRetrieved(26)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sCoCltID1") = ""
   Set p_oCoClt1 = loClient
   p_sCoCltNm1 = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getCoClient2(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getCoClient2"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .isLTOID = IFNull(p_oAppDrivr.getConfiguration("LTMS"), "0") = "1"
      If .InitClient() = False Then GoTo endProc
   End With
   
   With p_oBaseClas
      If lsValue <> "" Then
         If Trim(LCase(lsValue)) = Trim(LCase(p_sCoCltNm2)) Then GoTo endProc
         lbExist = loClient.SearchClient(lsValue, False)
      Else
         GoTo endWithClear
      End If

      If Not lbExist Then
         lasName = GetSplitedName(lsValue)
         loClient.Master("sLastName") = lasName(0)
         loClient.Master("sFrstName") = lasName(1)
      End If

      If loClient.getClient Then
         Set p_oCoClt2 = loClient
      End If
      
      p_oCoMkrInfo2.sLastName = p_oCoClt2.Master("sLastName")
      p_oCoMkrInfo2.sFrstName = p_oCoClt2.Master("sFrstName")
      p_oCoMkrInfo2.sMiddName = p_oCoClt2.Master("sMiddName")
      p_oCoMkrInfo2.sSuffixNm = p_oCoClt2.Master("sSuffixNm")
      p_oCoMkrInfo2.sBirthDte = p_oCoClt2.Master("dBirthDte")
      p_oCoMkrInfo2.sBirthPlc = p_oCoClt2.Master("sBrthTown")
      p_oCoMkrInfo2.sHouseNox = p_oCoClt2.Master("sHouseNox")
      p_oCoMkrInfo2.sAddresxx = p_oCoClt2.Master("sAddressx")
      p_oCoMkrInfo2.sBarangay = p_oCoClt2.Master("sBrgyName")
      p_oCoMkrInfo2.sTownxxxx = p_oCoClt2.Master("xTownName")
      p_oCoMkrInfo2.sProvince = p_oCoClt2.Master("xProvName")
      p_oCoMkrInfo2.sMobileNo = p_oCoClt2.Master("sMobileNo")
      p_oCoMkrInfo2.sEmailAdd = p_oCoClt2.Master("sEmailAdd")
      p_oCoMkrInfo2.sLTOIDxxx = p_oCoClt2.Master("sLTOIDxxx")
      p_oCoMkrInfo2.sTaxIDNox = p_oCoClt2.Master("sTaxIDNox")
      p_oCoMkrInfo2.sFatherNm = p_oCoClt2.Master("sFatherID")
      p_oCoMkrInfo2.sMotherNm = p_oCoClt2.Master("sMotherID")
      p_oCoMkrInfo2.sCitizenx = p_oCoClt2.Master(84)
      p_oCoMkrInfo2.sGenderCd = p_oCoClt2.Master(4)
      p_oCoMkrInfo2.sOccptnNm = p_oCoClt2.Master(86)

      .Master("sCoCltID2") = p_oCoClt2.Master("sClientID")
      p_oCoClt2.Master("cMCClient") = "1"
      p_sCoCltNm2 = p_oCoClt2.Master("sLastName") + ", " + p_oCoClt2.Master("sFrstName") + " " + Trim(p_oCoClt2.Master("sSuffixNm")) + IIf(Trim(p_oCoClt2.Master("sSuffixNm")) = "", "", " ") + p_oCoClt2.Master("sMiddName")
   End With
   
   getCoClient2 = True
   
endProc:
   RaiseEvent MasterRetrieved(27)
   Exit Function
endWithClear:
   p_oBaseClas.Master("sCoCltID2") = ""
   Set p_oCoClt2 = loClient
   p_sCoCltNm2 = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

'@@@
Private Function getApplication() As Boolean
   Dim lsOldProc As String
   Dim lasMaster() As String
   Dim lnReply As Integer
   
   lsOldProc = "GetApplication"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   With p_oBaseClas
      If .Master("sClientID") = Empty Then GoTo endWithClear
                  
      'This transaction cannot be an MC Sales so set application to none
      If .Detail(0, "sSerialID") = Empty Then
         .Master("sApplicNo") = "N-O-N-E"
      ElseIf IsPromoRoute() Then
         .Master("sApplicNo") = "N-O-N-E"
      ' XerSys - 2014-09-02
      '  Include the following code in this conditional testing
'      End If
'
'      If .Master("sApplicNo") <> "N-O-N-E" Then
      Else
         .Master("sApplicNo") = p_oLRApplic.getApplication(.Master("sClientID"), .Master("dTransact"))
               
         If .Master("sApplicNo") = Empty Then
            If IsAutoQM() = False Or .Master("cPaymForm") = "1" Then
               lnReply = MsgBox("Customer Has no Credit Application!!!" & vbCrLf & _
                                 " Click [YES] to continue anyway," & vbCrLf & _
                                 " [No] to abort entry", vbCritical + vbYesNo, "Confirm")
               
               If lnReply = vbYes Then
                  'We don't have a choice but to reinitialize the p_oLRApplic object
                  .Master("sApplicNo") = "N-O-N-E"
               End If
            End If
         Else
            If p_oLRApplic.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
         End If
      End If
   End With
         
   RaiseEvent MasterRetrieved(10)
   
   getApplication = True
   
endProc:
   Exit Function
endWithClear:
   p_oBaseClas.Master("sApplicNo") = ""
   RaiseEvent MasterRetrieved(10)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

'@@@
Private Function getHotItemApplication() As Boolean
   Dim lsOldProc As String
   Dim loHIMApplication As ggcLRApplication.clsHIMApplication
   Dim lnCtr As Integer
   
   lsOldProc = "GetHotItemApplication"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   'On Error GoTo errProc
   
   With p_oBaseClas
      If .Master("sClientID") = Empty Then GoTo endWithClear
      Set loHIMApplication = New ggcLRApplication.clsHIMApplication
      With loHIMApplication
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         .Parent = pxeMODULENAME
         If .InitTransaction() = False Then GoTo endProc
      End With
         
      For lnCtr = 0 To p_oBaseClas.ItemCount() - 1
         If p_oBaseClas.Detail(lnCtr, "cHotItemx") = xeYes Then
            .Master("sApplicNo") = loHIMApplication.getApplication(.Master("sClientID"), .Master("dTransact"), .Detail(lnCtr, "sModelIDx"))
         End If
      Next
            
      If .Master("sApplicNo") = Empty Then
         MsgBox "Customer has no valid Hot Item Application!!!" & vbCrLf & _
                  "Please verify your entry then try again!!!", vbCritical, "WARNING"
         GoTo endProc
      Else
         If loHIMApplication.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
      End If
   End With
         
   RaiseEvent MasterRetrieved(10)
   
   getHotItemApplication = loHIMApplication.SelectTransaction()
   
endProc:
   Exit Function
endWithClear:
   p_oBaseClas.Master("sApplicNo") = ""
   RaiseEvent MasterRetrieved(10)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & " )"
   GoTo endProc
End Function

Private Function getTerm(ByVal Value As String, ByVal Search As Boolean) As Boolean
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "GetTerm"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
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
   
      lsMaster = .GetMaster(22, lsMaster)
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
   RaiseEvent MasterRetrieved(5)
   Exit Function
endWithClear:
   With p_oBaseClas
      .Master("sTermIDxx") = ""
      .Master("sTermName") = ""
      .Master("dDueDatex") = .Master("dTransact")
   End With
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & Value _
                        & ", " & Search & " )"
End Function

Private Function getDetail(ByVal Row As Integer, _
                        ByVal Value As String, _
                        ByVal Search As Boolean) As Boolean
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String

   lsProcName = "getDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
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
      Debug.Print lsDetail
      If lsDetail = Empty Then
         If Search = False Then
            GoTo endWithClear
         Else
            GoTo endProc
         End If
      End If
      lasDetail = Split(lsDetail, "»")
      
      'Transfer
      If lasDetail(7) = Empty Then
         .Detail(Row, "cMotorNew") = xeYes
      Else
         .Detail(Row, "cMotorNew") = IIf(lasDetail(7) = xeNo, 1, 0)
      End If
      
      .Detail(Row, "cHotItemx") = IIf(.Detail(Row, "cMotorNew") = xeYes, lasDetail(11), 0)
      .Detail(Row, "cMotorUnt") = IIf(.Detail(Row, "cMotorNew") = xeYes, "0", "1")
      .Detail(Row, "cCSRValid") = IIf(.Detail(Row, "cMotorNew") = xeYes, lasDetail(12), 1)
      
      'kalyptus - 2020.06.29 12:03pm
      'Check for the cCSRValid of the Item Sold
      If Val(.Detail(Row, "cCSRValid")) = 0 Then
         lsDetail = p_oAppDrivr.getConfiguration("SOWoPNP")
         
         'Is sales of MCs without CSR not allowed
         If Val(lsDetail) = 0 Then
            MsgBox "Selling of MC(s) without CSR is not allowed!", vbCritical, "Warning"
            GoTo endWithClear
         'Is sales of MCs without CSR allowed through approval code
         ElseIf Val(lsDetail) = 2 Then
            MsgBox "Selling of MC(s) without CSR needs approval code!", vbCritical, "Warning"
         End If
      End If
      
      .Detail(Row, "sEngineNo") = lasDetail(1)
      .Detail(Row, "sFrameNox") = lasDetail(2)
      .Detail(Row, "sModelNme") = lasDetail(3)
      .Detail(Row, "sColorNme") = lasDetail(4)
      .Detail(Row, "sSerialID") = lasDetail(0)
      .Detail(Row, "nUnitPrce") = lasDetail(5)
      .Detail(Row, "sMCInvIDx") = lasDetail(6)
      .Detail(Row, "sReplMCID") = ""
      .Detail(Row, "nReplAmtx") = 0#
      .Detail(Row, "nReplIAmt") = 0#
      .Detail(Row, "xLedgerNo") = 0#
      .Detail(Row, "nLedgerNo") = lasDetail(8)
      .Detail(Row, "sModelIDx") = lasDetail(9)
      .Detail(Row, "sColorIDx") = lasDetail(10)
      
      
      ' our basis of receivable is always the first row
      If Row = 0 Then
         If .Detail(Row, "cMotorNew") = xeYes Then
            p_oMCPrice.MCSerialID = .Detail(Row, "sSerialID")
            p_oMCPrice.MCModelID = .Detail(Row, "sModelIDx")
            p_bValidate = p_oMCPrice.MCModelID <> ""
            
            If p_bValidate Then .Detail(Row, "nUnitPrce") = p_oMCPrice.SelPrice
         Else
            'mac 2020.07.31
            p_oMCPrice.MCSerialID = .Detail(Row, "sSerialID")
            p_bValidate = p_oMCPrice.MCModelID <> ""
                        
            If p_bValidate Then .Detail(Row, "nUnitPrce") = p_oMCPrice.SelPrice
         End If
      End If
      
      If .Master("cTranType") = 2 Then
         ' Replacement transaction must return the changed unit
         '  so get the mc info then process the returned unit
         If getReplacedUnit(Row) = False Then GoTo endWithClear
      End If
      
      ' call compute total
      Call computeTotal
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
      .Detail(Row, "sModelIDx") = ""
      .Detail(Row, "sColorIDx") = ""
      .Detail(Row, "cMotorUnt") = 0
      .Detail(Row, "cHotItemx") = 0
   
   End With
   RaiseEvent DetailRetrieved(1)
   RaiseEvent DetailRetrieved(2)
   RaiseEvent DetailRetrieved(3)
   RaiseEvent DetailRetrieved(4)
   RaiseEvent DetailRetrieved(7)
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & Row _
                        & ", " & Value _
                        & ", " & Search & " )"
End Function

Function getMAPFREReceipt() As String
   Dim lsSQL As String
   Dim lsProcName As String
   Dim loRS As Recordset
   
   lsProcName = "getDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   getMAPFREReceipt = ""
   lsSQL = " SELECT sTransNox" & _
            " FROM Receipt_Master" & _
            " WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
               " AND sSystemCd = 'MC'" & _
               " AND sSourceCd = 'MCSl'" & _
               " AND cTranType = '5'" & _
               " AND sInsTypID = '0109'" & _
               " AND cRegisFrm = '6'"
               
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If loRS.RecordCount = 1 Then
      getMAPFREReceipt = loRS("sTransNox")
   End If
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName
   GoTo endProc
End Function

Private Function getInsurance(ByVal Row As Integer, _
                           ByVal Value As String, _
                           ByVal Search As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsProcName As String
   Dim lsSQL As String

   lsProcName = "getDetail"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  sInsTypID" & _
               ", sInsTypNm" & _
               ", cRegisTyp" & _
            " FROM Insurance_Type" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive)
   
   If Value <> "" Then
      If p_oBaseClas.Detail(Row, "sInsTypNm") = Value Then GoTo endProc
      If Search Then
         lsSQL = AddCondition(lsSQL, "sInsTypNm LIKE " & strParm(Trim(Value) & "%"))
      Else
         lsSQL = AddCondition(lsSQL, "sInsTypNm = " & strParm(Trim(Value)))
      End If
   ElseIf Search = False Then
      GoTo endWithClear
   End If

   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, , adCmdText
   
   If loRS.EOF Then
      If Search = False Then
         GoTo endWithClear
      Else
         GoTo endProc
      End If
   End If
   
   With p_oBaseClas
      If loRS.RecordCount = 1 Then
         .Detail(Row, "sInsTypID") = loRS("sInsTypID")
         .Detail(Row, "sInsTypNm") = loRS("sInsTypNm")
         .Detail(Row, "cRegisTyp") = loRS("cRegisTyp")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, loRS, "sInsTypNm", "Insurance Type")
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
      End If
   End With
   
   Debug.Print p_oBaseClas.Detail(Row, "sInsTypNm")
   
   RaiseEvent DetailRetrieved(5)
   
   'Get Form of Registration Processing
   If p_oBaseClas.Detail(Row, "sInsTypNm") = "" Then
      p_oBaseClas.Detail(Row, "cMotorUnt") = "0"
   ElseIf p_oBaseClas.Detail(Row, "cMotorNew") = xeYes Then
      p_oBaseClas.Detail(Row, "cMotorUnt") = "0"
   Else
      Set frmRegisProcess.AppDriver = p_oAppDrivr
      Load frmRegisProcess
      With frmRegisProcess
         .TransNo = p_oBaseClas.Master("sTransNox")
         .EngineNo = p_oBaseClas.Detail(Row, "sEngineNo")
         .FrameNo = p_oBaseClas.Detail(Row, "sFrameNox")
         .ModelName = p_oBaseClas.Detail(Row, "sModelNme")
         Call .setRegisForm(Val(p_oBaseClas.Detail(Row, "cMotorUnt")))
         .Show 1
         p_oBaseClas.Detail(Row, "cMotorUnt") = .getRegisForm
      End With
      Unload frmRegisProcess
   End If
   
   getInsurance = True
   
endProc:
   Set loRS = Nothing
   Exit Function
endWithClear:
   With p_oBaseClas
      .Detail(Row, "sInsTypID") = ""
      .Detail(Row, "sInsTypNm") = ""
      .Detail(Row, "cRegisTyp") = ""
      .Detail(Row, "cMotorUnt") = "0"
   End With
   RaiseEvent DetailRetrieved(5)
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & Row _
                        & ", " & Value _
                        & ", " & Search & " )"
   GoTo endProc
End Function

Private Function delOtherTrans() As Boolean
   Dim loOReceipt As clsORReceipt
   Dim loCltTrans As clsClientTrans
   Dim loFinTrans As clsFinanceMaster
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsClientID As String
   
   lsProcName = "delOtherTrans"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   With p_oBaseClas
      If Not delMCTransaction() Then GoTo endProc
      
      Set loRS = New Recordset
      lsSQL = AddCondition(.DetailQuery, _
                  "a.sTransNox = " & strParm(.Master("sTransNox")))
      Debug.Print lsSQL
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If loRS.EOF Then GoTo endProc
      
      ' check if sales has amount paid
      If .Master("nAmtPaidx") > 0 Then
         Set loOReceipt = New clsORReceipt
         Set loOReceipt.AppDriver = p_oAppDrivr
         loOReceipt.Branch = p_sBranchCd
         loOReceipt.Parent = "clsMCSales"
               
         lsClientID = p_oBaseClas.Master("sClientID")
         loOReceipt.Filter = "a.sClientID = " & strParm(lsClientID)
         
         If loOReceipt.InitTransaction = False Then GoTo endProc
         If loOReceipt.SearchByRef(p_oBaseClas.Master("sTransNox"), "MCSl") = False Then
            MsgBox "Unable to Load Receipt Info!", vbCritical, "Warning"
            GoTo endProc
         End If
   
         'We should be specific on what we really want to do
         'kalyptus - 2011-08-31
         If p_bIsDelete Then
            If delAssuredInfo = False Then GoTo endProc
            If loOReceipt.DeleteTransaction = False Then GoTo endProc
         Else
            If cancelAssuredInfo = False Then GoTo endProc
            If loOReceipt.CancelTransaction = False Then GoTo endProc
         End If
      End If
      
      '@@@@
      If p_sApplicNo <> Empty Then
         If p_oLRMaster.DeleteAccount = False Then GoTo endProc
         If .MasterOriginal("sApplicNo") <> "N-O-N-E" Then
            If .MasterOriginal("cPaymForm") <> "0" Then
               ' XerSys 2014-05-30
               '  Disable the following code, and just update the status of the
               '     credit application to approved
               
               lsSQL = "UPDATE MC_Credit_Application" & _
                        " SET cTranStat = " & strParm(xeStatePosted) & _
                           ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
                        " WHERE sTransNox = " & strParm(.Master("sApplicNo"))
               If p_oAppDrivr.Execute(lsSQL, "MC_Credit_Application") <= 0 Then
                  MsgBox "Unable to update credit application info!" & vbCrLf & _
                           "Please inform GGC SEG about this matter!", vbCritical, "Warning"
                  GoTo endProc
               End If
   '            ' add this condition to test if a record was loaded or not
   '            If p_oLRApplic.EditMode = xeModeUnknown Then
   '               If p_oLRApplic.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
   '               If p_oLRApplic.DeleteTransaction = False Then GoTo endProc
   '            End If
            End If
         End If
      End If
      
      'All term payment form should have a ledger
      If .Master("cPaymForm") = "3" Or .Master("cPaymForm") = "4" Then
         'Delete from Customer Ledger
            If SaveMCCustomerTrans(xeModeDelete) = False Then GoTo endProc
      ElseIf .Master("cPaymForm") = "5" Then
         'Delete from Finance Master
         Set loFinTrans = New clsFinanceMaster
         Set loFinTrans.AppDriver = p_oAppDrivr
         loFinTrans.Branch = p_sBranchCd
         loFinTrans.Parent = pxeMODULENAME
         Call loFinTrans.InitAccount
      
         If loFinTrans.LoadByCondition("sReferNox = " & strParm(p_oBaseClas.Master("sTransNox"))) = True Then
            loFinTrans.DeleteAccount
         End If
      End If
      
      If pbHasGAway Then
               'We should be specific on what we really want to do
               'kalyptus - 2011-08-31
         If p_bIsDelete Then
            If p_oGiveAway.DeleteTransaction() = False Then GoTo endProc
         Else
            If p_oGiveAway.CancelTransaction() = False Then GoTo endProc
         End If
      End If
   
      'kalyptus - 2017.05.09 03:58pm
      'Remove auto-deletion of GCard Application record and replace with a utility
      'from the main office...
'      lsSQL = "DELETE FROM G_Card_Application" & _
'               " WHERE sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode + "%") & _
'               " AND sSourceCd = 'M02910000005' " & _
'               " AND sSourceNo = " & strParm(.Master("sDRNoxxxx"))
'      Call p_oAppDrivr.Execute(lsSQL, "G_Card_Application")
   End With
      
   delOtherTrans = True
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function saveReceipt(lcTranStat As xeTransactionStatus, _
                                 lnTranAmtx As Double) As Boolean
   Dim loOReceipt As clsORReceipt
   Dim lsProcName As String
   Dim lsClientID As String
   
   lsProcName = "saveReceipt"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If Trim(p_sPRNoxxxx) = "" And _
      Trim(p_oBaseClas.Master("sORNoxxxx")) = "" Then
      If p_oBaseClas.Master("cPaymForm") = "3" Or p_oBaseClas.Master("cPaymForm") = "4" Then
         saveReceipt = True
      End If
      
      If p_nAdvPaymx = p_oBaseClas.Master("nTranTotl") Then
         saveReceipt = True
      End If
      GoTo endProc
   End If
   
   Set loOReceipt = New clsORReceipt
   With loOReceipt
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .Parent = "clsMCSales"
            
      lsClientID = IIf(IsNull(p_oBaseClas.MasterOriginal("sClientID")), p_oBaseClas.Master("sClientID"), p_oBaseClas.MasterOriginal("sClientID"))
      'kalyptus - 2015.11.17 11:13am
      'Include cTranType to the filter
      .Filter = "a.sClientID = " & strParm(lsClientID) & " AND a.cTranType = '0'"
      
      If .InitTransaction = False Then GoTo endProc
   
      If p_oBaseClas.EditMode = xeModeAddNew Then
         If .NewTransaction = False Then GoTo endProc
      Else
         If .SearchByRef(p_oBaseClas.Master("sTransNox"), "MCSl") = False Then
            If .NewTransaction = False Then GoTo endProc
         Else
            .UpdateTransaction
         End If
      End If
      
      'kalyptus - 2015.06.19 09:09am
      'Incorporate gift coupon here...
      Set loOReceipt.GiftCoupon = p_oGift
      Debug.Print loOReceipt.GiftCoupon.Detail(1, "gcertnox")
      
      .Master("dTransact") = p_oBaseClas.Master("dTransact")
      .Master("sORNoxxxx") = p_oBaseClas.Master("sORNoxxxx")
      .Master("sRemarksx") = p_oBaseClas.Master("sRemarksx")
      .Master("nTranAmtx") = lnTranAmtx - p_xChecks.Amountxx
      .Master("cTranType") = "0"
      .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
      .Master("sSourceCd") = "MCSl"
      .Master("sClientID") = p_oBaseClas.Master("sClientID")
      .Master("sCoCltID1") = p_oBaseClas.Master("sCoCltID1")
      .Master("sCoCltID2") = p_oBaseClas.Master("sCoCltID2")
      .Master("sSystemCD") = "MC"
      .Master("sSerialID") = p_oBaseClas.Detail(0, "sSerialID")
      .Master("sInsTypID") = p_oBaseClas.Detail(0, "sInsTypID")
      .Checks("sCheckNox") = p_xChecks.CheckNox
      .Checks("sAcctNoxx") = p_xChecks.AcctNoxx
      .Checks("sBankIDxx") = p_xChecks.BankIDxx
      .Checks("dCheckDte") = p_xChecks.CheckDte
      .Checks("nClearing") = p_xChecks.Clearing
      .Checks("nAmountxx") = p_xChecks.Amountxx
      .Checks("bIssuedOR") = p_bIssuedOR
      .Master("cTranStat") = lcTranStat
      .PRNumber = p_sPRNoxxxx
      
      If .SaveTransaction = False Then GoTo endProc
      
      ' after saving the receipt, save the transaction ledger
      If p_oBaseClas.Master("sApplicNo") <> Empty And _
         p_oBaseClas.Master("cPaymForm") <> "0" Then
         p_oLRMaster.Master("sORNoxxxx") = .Master("sORNoxxxx")
         
         If p_oBaseClas.Detail(0, "cMotorNew") = "1" Then ' Brand New
            p_oLRMaster.Master("nDownTotl") = (p_oBaseClas.Master("nAmtPaidx") + p_nAdvPaymx) - _
                                                IIf(p_oLRMaster.Master("nAcctTerm") = 0, 0, p_nChattelM) - p_nCommAmtx
         Else
            p_oLRMaster.Master("nDownTotl") = (p_oBaseClas.Master("nAmtPaidx") + p_nAdvPaymx) - _
                                                IIf(p_oLRMaster.Master("nAcctTerm") = 0, 0, p_oBaseClas.Detail(0, "nInsAmtxx")) - p_nCommAmtx
         End If
         p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nABalance") - p_oLRMaster.Master("nDownTotl")
      End If
   End With
   
   saveReceipt = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

' XerSys 2014-03-25
'  Only update the MC Customer Order object
Private Function updateAdvPayment() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loMCOrder As clsMCCustOrder
   
   Dim lsTransNox As String
   Dim lnCashAmtx As Double
   Dim lnCheckAmt As Double
   
   lsProcName = "updateAdvPayment"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loMCOrder = New clsMCCustOrder
   With loMCOrder
      Set .AppDriver = p_oAppDrivr
      .Parent = pxeMODULENAME
      If Not .InitTransaction Then GoTo endProc
   End With
   
   With p_oAdvPaym
      .MoveFirst
      
      lsTransNox = .Fields("sTransNox")
      lnCashAmtx = 0
      lnCheckAmt = 0
      
      Do Until .EOF()
         If lsTransNox <> .Fields("sTransNox") Then
            If lnCashAmtx + lnCheckAmt > 0 Then
               If Not loMCOrder.CreditAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt, Master("dTransact")) Then
                  GoTo endProc
               End If
               If lnCashAmtx > 0 Then
                  lsSQL = "INSERT INTO MC_SO_Reservation" & _
                           " SET sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                              ", sReferNox = " & strParm(lsTransNox) & _
                              ", sSourceCd = " & strParm(xeSrcOReceipt) & _
                              ", nAmountxx = " & lnCashAmtx
                  If p_oAppDrivr.Execute(lsSQL, "MC_SO_Reservation") < 0 Then
                     MsgBox "Unable to Insert Sales Reservation Info!" & vbCrLf & _
                           "Please inform Guanzon's SSG/SEG about this matter", vbCritical, "Warning"
                     GoTo endProc
                  End If
               End If
               If lnCheckAmt > 0 Then
                  lsSQL = "INSERT INTO MC_SO_Reservation" & _
                           " SET sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                              ", sReferNox = " & strParm(lsTransNox) & _
                              ", sSourceCd = " & strParm(xeSrcPReceipt) & _
                              ", nAmountxx = " & lnCheckAmt
                  If p_oAppDrivr.Execute(lsSQL, "MC_SO_Reservation") < 0 Then
                     MsgBox "Unable to Insert Sales Reservation Info!" & vbCrLf & _
                           "Please inform Guanzon's SSG/SEG about this matter", vbCritical, "Warning"
                     GoTo endProc
                  End If
               End If
            End If
            lsTransNox = .Fields("sTransNox")
            lnCashAmtx = 0
            lnCheckAmt = 0
         End If
         
         If .Fields("xSrcTrans") = xeSrcOReceipt Then
            lnCashAmtx = .Fields("nCredtAmt")
         Else
            lnCheckAmt = .Fields("nCredtAmt")
         End If
         
         .MoveNext
      Loop
      
      If lnCashAmtx + lnCheckAmt > 0 Then
         If Not loMCOrder.CreditAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt, Master("dTransact")) Then
            GoTo endProc
         End If
         
         If lnCashAmtx > 0 Then
            lsSQL = "INSERT INTO MC_SO_Reservation" & _
                     " SET sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                        ", sReferNox = " & strParm(lsTransNox) & _
                        ", sSourceCd = " & strParm(xeSrcOReceipt) & _
                        ", nAmountxx = " & lnCashAmtx
            If p_oAppDrivr.Execute(lsSQL, "MC_SO_Reservation") < 0 Then
               MsgBox "Unable to Insert Sales Reservation Info!" & vbCrLf & _
                     "Please inform Guanzon's SSG/SEG about this matter", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         If lnCheckAmt > 0 Then
            lsSQL = "INSERT INTO MC_SO_Reservation" & _
                     " SET sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                        ", sReferNox = " & strParm(lsTransNox) & _
                        ", sSourceCd = " & strParm(xeSrcPReceipt) & _
                        ", nAmountxx = " & lnCheckAmt
            If p_oAppDrivr.Execute(lsSQL, "MC_SO_Reservation") < 0 Then
               MsgBox "Unable to Insert Sales Reservation Info!" & vbCrLf & _
                     "Please inform Guanzon's SSG/SEG about this matter", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
   End With

   ' XerSys 2014-03-25
   '  Disable the following codes
'   Dim loPReceipt As clsPRReceipt
'   Dim loOReceipt As clsORReceipt
'   Dim loChecks As clsCheckClearing
'   Set loPReceipt = New clsPRReceipt
'   With loPReceipt
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .Parent = pxeMODULENAME
'
'      If .InitTransaction = False Then GoTo endProc
'
'      ' first update the selected advance payment
'      lasRefer = Split(p_sAdvRefer, "»")
'
'      ' XerSys - 2014-01-27
'      '  Check if advance payment came from PR or OR
'      For lnCtr = 0 To UBound(lasRefer)
'         If Right(lasRefer(lnCtr), 2) = "PR" Then
'            If InStr(1, p_sAdvRefB4, lasRefer(lnCtr)) = 0 Then
'               If .OpenTransaction(Left(lasRefer(lnCtr), pxeTransNox)) = False Then GoTo endProc
'               Call .UpdateTransaction
'               If .Checks("nAmountxx") > 0 Then
'                  .Master("cTranStat") = xeStatePosted
'               Else
'                  .Master("cTranStat") = xeStateClosed
'               End If
'               .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
'               .Master("sSourceCd") = "MCSl"
'               If .SaveTransaction() = False Then GoTo endProc
'
'               'Check if this is an advance payment using a check...
'               If .Checks("nAmountxx") > 0 Then
'                  Set loChecks = New clsCheckClearing
'                  Set loChecks.AppDriver = p_oAppDrivr
'                  loChecks.Branch = p_sBranchCd
'                  loChecks.Parent = pxeMODULENAME
'                  If loChecks.InitTransaction = False Then GoTo endProc
'                  If loChecks.SearchTransaction(.Master("sTransNox"), True) = False Then GoTo endProc
'                  loChecks.UpdateTransaction
'                  loChecks.Master("cChckStat") = 1 'xeClearedCheck
'                  loChecks.Master("dStatChng") = p_oBaseClas.Master("dTransact")
'                  If loChecks.SaveTransaction = False Then GoTo endProc
'               End If
'            End If
'         End If
'      Next
'
'      ' then undo previously selected advance payment
'      lasRefer = Split(p_sAdvRefB4, "»")
'      For lnCtr = 0 To UBound(lasRefer)
'         If Right(lasRefer(lnCtr), 2) = "PR" Then
'            If InStr(1, p_sAdvRefer, lasRefer(lnCtr)) = 0 Then
'               If .OpenTransaction(Left(lasRefer(lnCtr), pxeTransNox)) = False Then GoTo endProc
'
'               Call .UpdateTransaction
'               .Master("cTranStat") = xeStateOpen
'               .Master("sReferNox") = ""
'               .Master("sSourceCd") = ""
'               If .SaveTransaction() = False Then GoTo endProc
'
'               'Check if this is an advance payment using a check...
'               If .Checks("nAmountxx") > 0 Then
'                  Set loChecks = New clsCheckClearing
'                  Set loChecks.AppDriver = p_oAppDrivr
'                  loChecks.Branch = p_sBranchCd
'                  loChecks.Parent = pxeMODULENAME
'                  If loChecks.InitTransaction = False Then GoTo endProc
'                  If loChecks.SearchTransaction(.Master("sTransNox"), True) = False Then GoTo endProc
'                  loChecks.UpdateTransaction
'                  loChecks.Master("cChckStat") = 0 'xeOpenCheck
'                  loChecks.Master("dStatChng") = .Master("dTransact") 'Save the previous date from PR
'                  If loChecks.SaveTransaction = False Then GoTo endProc
'               End If
'            End If
'         End If
'      Next
'   End With
'
'   ' XerSys - 2014-01-27
'   '  Update reservation paid in cash
'   Set loOReceipt = New clsORReceipt
'   With loOReceipt
'      Set .AppDriver = p_oAppDrivr
'      .Branch = p_sBranchCd
'      .Parent = pxeMODULENAME
'
'      If .InitTransaction = False Then GoTo endProc
'
'      ' first update the selected advance payment
'      lasRefer = Split(p_sAdvRefer, "»")
'
'      ' XerSys - 2014-01-27
'      '  Only advance payment paid by cash must be proccessed
'      For lnCtr = 0 To UBound(lasRefer)
'         If Right(lasRefer(lnCtr), 2) = "OR" Then
'            If InStr(1, p_sAdvRefB4, lasRefer(lnCtr)) = 0 Then
'               If .OpenTransaction(Left(lasRefer(lnCtr), pxeTransNox)) = False Then GoTo endProc
'
'               Call .UpdateTransaction
'               .Master("cTranStat") = xeStatePosted
'               .Master("sReferNox") = p_oBaseClas.Master("sTransNox")
'               .Master("sSourceCd") = "MCSl"
'               If .SaveTransaction() = False Then GoTo endProc
'            End If
'         End If
'      Next
'
'      ' then undo previously selected advance payment
'      lasRefer = Split(p_sAdvRefB4, "»")
'      For lnCtr = 0 To UBound(lasRefer)
'         If Right(lasRefer(lnCtr), 2) = "PR" Then
'            If InStr(1, p_sAdvRefer, lasRefer(lnCtr)) = 0 Then
'               If .OpenTransaction(Left(lasRefer(lnCtr), pxeTransNox)) = False Then GoTo endProc
'
'               Call .UpdateTransaction
'               .Master("cTranStat") = xeStateClosed
'               .Master("sReferNox") = ""
'               .Master("sSourceCd") = ""
'               If .SaveTransaction() = False Then GoTo endProc
'            End If
'         End If
'      Next
'   End With

   updateAdvPayment = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

'Use in saving mc units purchase
Private Function saveMCTransaction() As Boolean
   Dim loMCTrans As clsMCSerialTrans
   Dim loMCReturned As clsMCSerialTrans
   Dim loMCRegistration As clsSalesRegistration
   Dim lsProcName As String
   Dim lnRow As Integer, lnInsurance As Integer
   Dim lnRegis As Integer
   Dim lsClients As String
   Dim lbWithReg As Boolean
   
   lsProcName = "saveMCTransaction"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loMCTrans = New clsMCSerialTrans
   With loMCTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction(pxeMCSales) = False Then GoTo endProc
   End With
   
   If p_oBaseClas.Master("cTranType") = "2" Then
      Set loMCReturned = New clsMCSerialTrans
      With loMCReturned
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         If .InitTransaction(pxeMCSalesReturn) = False Then GoTo endProc
      End With
   End If
   
   Set loMCRegistration = New clsSalesRegistration
   With loMCRegistration
      Set .AppDriver = p_oAppDrivr
      If .InitTransaction() = False Then GoTo endProc
      .PaymForm = p_oBaseClas.Master("cPaymForm")
      .TranType = p_oBaseClas.Master("cTranType")
   End With
   
   lnInsurance = vbCancel
   With p_oBaseClas
      lnRow = 0
      lnRegis = 0
      lbWithReg = False
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sEngineNo") = "" Then Exit For
         
         loMCTrans.Detail(pnCtr, "sSerialID") = .Detail(pnCtr, "sSerialID")
         loMCTrans.Detail(pnCtr, "sMCInvIDx") = .Detail(pnCtr, "sMCInvIDx")
         loMCTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         loMCTrans.Detail(pnCtr, "cSoldStat") = IIf(.Detail(pnCtr, "cMotorNew") = xeYes, 0, 1)
         loMCTrans.Detail(pnCtr, "sWarrntNo") = .Detail(pnCtr, "sWarrntNo")
         
         If .Master("cTranType") = "2" Then
            If .Detail(pnCtr, "sSerialID") = Empty Then
               MsgBox "Please select the Motorcycle that needs to be replaced!!!" & vbCrLf & _
                      "Please inform Guanzon's SSG/SEG if you have question.", vbCritical, "Warning"
               GoTo endProc
            End If
            loMCReturned.Detail(lnRow, "sSerialID") = .Detail(pnCtr, "sReplMCID")
            loMCReturned.Detail(lnRow, "sMCInvIDx") = .Detail(pnCtr, "sRMCInvID")
            loMCReturned.Detail(lnRow, "nLedgerNo") = .Detail(pnCtr, "xLedgerNo")
            loMCReturned.Detail(lnRow, "cSoldStat") = 1
            lnRow = lnRow + 1
         End If
                  
         loMCRegistration.Detail(pnCtr + 1, "sSerialID") = .Detail(pnCtr, "sSerialID")
         loMCRegistration.Detail(pnCtr + 1, "sInsTypID") = .Detail(pnCtr, "sInsTypID")
         loMCRegistration.Detail(pnCtr + 1, "cRegisTyp") = .Detail(pnCtr, "cRegisTyp")
         loMCRegistration.Detail(pnCtr + 1, "nInsAmtxx") = .Detail(pnCtr, "nInsAmtxx")
         loMCRegistration.Detail(pnCtr + 1, "cMotorNew") = .Detail(pnCtr, "cMotorUnt") + 1
         loMCRegistration.Detail(pnCtr + 1, "sReplMCID") = .Detail(pnCtr, "sReplMCID")
         loMCRegistration.Detail(pnCtr + 1, "nReplIAmt") = .Detail(pnCtr, "nReplIAmt")

         If .EditMode = xeModeUpdate Then
            loMCRegistration.Detail(pnCtr + 1, "sPrevSerl") = .DetailOriginal(pnCtr, "sSerialID")
         End If
                  
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
         
'660            If lnInsurance = vbCancel Then
'670               If .Detail(pnCtr, 5) = Empty And .Detail(pnCtr, "cMotorNew") = 1 Then
'680                  lnRow = MsgBox("Some of the Motorcycles Sold Does Not Have Registration!!!" & vbCrLf & _
'                              "    Continue Saving Anyway?", vbQuestion + vbYesNo, "Confirm")
'690                  If lnRow <> vbYes Then
'700                     GoTo endProc
'710                  Else
'720                     lnInsurance = vbYes
'730                  End If
'740               ElseIf .Detail(0, "cMotorNew") = xeNo And .Detail(0, "nInsAmtxx") = 0 Then
'750                  lnRow = MsgBox("Some of the motorcycle Sold Does Not Have Deed of Sale Payment!!!" & vbCrLf & _
'                              "    Continue Saving Anyway?", vbQuestion + vbYesNo, "Confirm")
'760                  If lnRow <> vbYes Then
'770                     GoTo endProc
'780                  Else
'790                     lnInsurance = vbYes
'800                  End If
'810               End If
'820            End If

               'Validate motorcycle registration and/or deed of sale
               'Is the mc sales a new motorcycle?
         If .Detail(pnCtr, "cMotorNew") = "1" Then
                  'Is mc has insurance?
            If .Detail(pnCtr, 5) = Empty Then
                     'Does it have an amount allocated for registration?
               If .Detail(pnCtr, "nInsAmtxx") = 0 Then
                  lnRow = MsgBox("Motorcycles with Engine " & .Detail(pnCtr, "sEngineNo") & "  Does Not Have Registration!!!" & vbCrLf & _
                                 "  Continue Saving Anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
                  ' XerSys - 2015-05-14
                  '  Ask marlon if the condition is correct. for now replace the operand.
                  If lnRow <> vbYes Then '=
                     GoTo endProc
                  End If
                     'If it has but empty then inform user about a possible problem in the entry
               Else
                  Call MsgBox("Motorcycles with Engine " & .Detail(pnCtr, "sEngineNo") & " has invalid registration information!!!" & vbCrLf & _
                                 " Please check your entry!!!", vbInformation + vbOKOnly, "Warning")
                  GoTo endProc
               End If
            End If
               'MC sales is a repo unit!
         Else
                  'Does it have an insurance amount?
            If .Detail(pnCtr, "nInsAmtxx") = 0 Then
               lnRow = MsgBox("Motorcycle with Engine " & .Detail(pnCtr, "sEngineNo") & "  Does Not Have Deed of Sale Payment!!!" & vbCrLf & _
                              " Continue Saving Anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
               ' XerSys - 2015-05-14
               '  Ask marlon if the condition is correct. for now replace the operand.
               If lnRow <> vbYes Then '=
                  GoTo endProc
               End If
            Else
                     'Check validity of amount allocated
               If .Detail(pnCtr, 5) = Empty Then
                        'Is allocation for deed of sale more than the minimum amount intended
                  If .Detail(pnCtr, "nInsAmtxx") > pxeMinReg_DED Then
                     lnRow = MsgBox("Amount Allocated for Motorcycle with Engine " & .Detail(pnCtr, "sEngineNo") & " seems to be for Registration!!!" & vbCrLf & _
                                    " Continue Saving Anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
                     ' XerSys - 2015-05-14
                     '  Ask marlon if the condition is correct. for now replace the operand.
                     If lnRow <> vbYes Then '=
                        GoTo endProc
                     End If
                  End If
               Else
                        'Is allocation for registration intended for Deed of Sale
                  If .Detail(pnCtr, "nInsAmtxx") <= pxeMinReg_DED Then
                     lnRow = MsgBox("Amount Allocated for Motorcycle with Engine " & .Detail(pnCtr, "sEngineNo") & " seems to be for Deed of Sale!!!" & vbCrLf & _
                                    " Continue Saving Anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
                     ' XerSys - 2015-05-14
                     '  Ask marlon if the condition is correct. for now replace the operand.
                     If lnRow <> vbYes Then '=
                        GoTo endProc
                     End If
                  Else
                     'kalyptus - 2015.08.01 01:35pm
                     'This will allocate zero(0) downpayment customers to have MC_Registration Records....
                     lbWithReg = True
                  End If
               End If
            End If
         End If

      Next
      
      lsClients = .Master("sClientID")
      lsClients = lsClients & IIf(IFNull(.Master("sCoCltID1")) = "", "", "»" & .Master("sCoCltID1"))
      lsClients = lsClients & IIf(IFNull(.Master("sCoCltID2")) = "", "", "»" & .Master("sCoCltID2"))
      
      If loMCTrans.Detail(0, "sSerialID") <> Empty Then
         If .Master("cTranType") = "2" Then
            If Not loMCTrans.Sales(.Master("sTransNox") _
                        , .Master("dTransact") _
                        , lsClients _
                        , .EditMode) Then
               GoTo endProc
            End If
            If Not loMCReturned.SalesReturn(.Master("sTransNox") _
                        , .Master("dTransact") _
                        , .EditMode) Then
               GoTo endProc
            End If
         Else
            If Not loMCTrans.Sales(.Master("sTransNox") _
                        , .Master("dTransact") _
                        , lsClients _
                        , .EditMode) Then
               GoTo endProc
            End If
         End If
            
         'Kalyptus - 2014.09.06 09:33am
         'Include p_nAdvPaymx in our condition
         '   - Since there are instances that we haven't create a Registration Record if OR is empty...
         If .Master("sORNoxxxx") <> "" Or _
            .Master("cPaymForm") = "4" Or _
            .Master("cTranType") = "2" Or _
            p_nAdvPaymx > 0 Or _
            lbWithReg Then
            'Save registration if any of this condition is true
            '1. Cash Payment
            '2. Term With Registration(for checks)
            '3. Replacement with previous Registration
            '4. Repo Sales with zero(0) downpayment but with insurance and insurance amount entry
            If Not loMCRegistration.SaveTransaction(.Master("sTransNox") _
                        , .Master("dTransact") _
                        , lsClients _
                        , .EditMode) Then
               GoTo endProc
            End If
         Else
            If .MasterOriginal("sORNoxxxx") <> "" Then
               If Not loMCRegistration.SaveTransaction(.Master("sTransNox") _
                           , .Master("dTransact") _
                           , lsClients _
                           , xeModeDelete) Then
                  GoTo endProc
               End If
            End If
         End If
      End If
   End With
   
   saveMCTransaction = True
   
endProc:
   Set loMCTrans = Nothing
   Set loMCReturned = Nothing
   Set loMCRegistration = Nothing
   
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
   GoTo endProc
End Function

Private Function getReplacedUnit(ByVal Row As Integer) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lasSQL() As String
   
   lsProcName = "getReplacedUnit(" & Row & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

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
                  ", a.nLedgerNo" & _
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
      
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   ' check if there's unit bought by the customer
   If loRS.EOF Then
      MsgBox "Customer Did Not Bought Motorcycle Here!!!" & _
               "Motorcycle Replacement is Not Allowed!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, loRS, _
                  "sEngineNo»sModelNme»nUnitPrce»nInsAmtxx", _
                  "Engine No»Model»Unit Price»Insurance")
      
      If lsSQL = Empty Then
         MsgBox "No Selection Was Made!!!" & _
                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      loRS.MoveFirst
      lasSQL = Split(lsSQL, "»")
      loRS.Find "sSerialID = " & strParm(lasSQL(0))
'      loRS.Find "sSerialID = " & strParm(Left(lsSQL, 10))
   End If
   
   With p_oBaseClas
      .Detail(Row, "sReplMCID") = loRS("sSerialID")
      .Detail(Row, "nReplAmtx") = loRS("nUnitPrce")
      .Detail(Row, "nReplIAmt") = loRS("nInsAmtxx")
      .Detail(Row, "sRMCInvID") = loRS("sMCInvIDx")
      .Detail(Row, "xLedgerNo") = loRS("nLedgerNo")
   
'      ' XerSys - 2014-07-01
'      '  Check if loan receivable exist
'      lsSQL = "SELECT sAcctNmbr" & _
'               " FROM MC_AR_Master" & _
'               " WHERE sClientID = " & strParm(.Master("sClientID")) & _
'                  " AND sSerialID = " & strParm(.Detail(Row, "sReplMCID"))
'      loRS.Close
'      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
'
'      If Not loRS.EOF Then
'         ' reload loan receivable
'         Call p_oLRMaster.OpenAccount(loRS("sAcctNmbr"))
'      End If
   End With
   getReplacedUnit = True
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName
   GoTo endProc
End Function

Private Function delMCTransaction() As Boolean
   Dim loMCTrans As clsMCSerialTrans
   Dim loMCRegistration As clsSalesRegistration
   Dim lsProcName As String
   
   lsProcName = "delMCTransaction()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Set loMCTrans = New clsMCSerialTrans
   With loMCTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitTransaction(pxeMCSales) = False Then
         MsgBox "MCTrans"
         GoTo endProc
      End If
   End With
   
   Set loMCRegistration = New clsSalesRegistration
   With loMCRegistration
      Set .AppDriver = p_oAppDrivr
      If .InitTransaction = False Then
         MsgBox "MCTrans"
         GoTo endProc
      End If
   End With
   
   With p_oBaseClas
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sSerialID") = Empty Then Exit For
         
         loMCTrans.Detail(pnCtr, "sSerialID") = .Detail(pnCtr, "sSerialID")
         loMCTrans.Detail(pnCtr, "sMCInvIDx") = .Detail(pnCtr, "sMCInvIDx")
         loMCTrans.Detail(pnCtr, "nLedgerNo") = .Detail(pnCtr, "nLedgerNo")
         loMCTrans.Detail(pnCtr, "cSoldStat") = xeNo
         
         loMCRegistration.Detail(pnCtr + 1, "sSerialID") = .Detail(pnCtr, "sSerialID")
         loMCRegistration.Detail(pnCtr + 1, "sInsTypID") = .Detail(pnCtr, "sInsTypID")
         loMCRegistration.Detail(pnCtr + 1, "cRegisTyp") = .Detail(pnCtr, "cRegisTyp")
         loMCRegistration.Detail(pnCtr + 1, "nInsAmtxx") = .Detail(pnCtr, "nInsAmtxx")
         loMCRegistration.Detail(pnCtr + 1, "cMotorNew") = .Detail(pnCtr, "cMotorNew")
      Next
      
      If Not loMCTrans.Sales(.Master("sTransNox"), _
               .Master("dTransact"), _
               .Master("sClientID"), _
               xeModeDelete) Then
         GoTo endProc
      End If
      
      If Not loMCRegistration.SaveTransaction(.Master("sTransNox") _
               , .Master("dTransact") _
               , .Master("sClientID") _
               , xeModeDelete) Then
         GoTo endProc
      End If
   End With
   
   delMCTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

Private Sub createAdvPaym()
   Set p_oAdvPaym = New Recordset
   With p_oAdvPaym
      .Fields.Append "sTransNox", adVarChar, 12
      .Fields.Append "sReferNox", adVarChar, 8
      .Fields.Append "dTransact", adDate
      .Fields.Append "nTranAmtx", adCurrency, 12
      .Fields.Append "nCredtAmt", adCurrency, 12
      .Fields.Append "cTranStat", adChar, 1
      .Fields.Append "sCheckTrn", adChar, 12
      .Fields.Append "cDepositd", adChar, 1
      .Fields.Append "cChckStat", adChar, 1
      .Fields.Append "xSrcTrans", adVarChar, 4
      .Open
   End With
End Sub

Private Function getAdvancePayment() As Boolean
   Dim loAdvPaym As frmAdvancePayment
   Dim loTmp As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   lsProcName = "getAdvancePayment"
   'On Error GoTo errProc
   
   If p_oBaseClas.Master("sClientID") = Empty Then GoTo endWithWarning
   
   ' XerSys - 2014-03-24
   '  All advance payment should be created through mc customer order
   lsSQL = "SELECT a.sTransNox" & _
               ", b.sPRNoxxxx sReferNox" & _
               ", a.dTransact" & _
               ", b.nTranAmtx - a.nCredtAmt nTranAmtx" & _
               ", a.nCredtChk nCredtAmt" & _
               ", b.cTranStat" & _
               ", IFNULL(c.sTransNox, '') xCheckTrn" & _
               ", IFNULL(c.cDepositd, '1') cDepositd" & _
               ", IFNULL(c.cChckStat, '2') cChckStat" & _
               ", " & strParm(xeSrcPReceipt) & " xSrcTrans" & _
            " FROM MC_Customer_Order a" & _
               ", Provisionary_Receipt_Master b" & _
                  " LEFT JOIN Checks_Received c" & _
                     " ON b.sTransNox = c.sReferNox" & _
                        " AND c.sSourceCd = " & strParm(xeSrcPReceipt) & _
            " WHERE a.sTransNox = b.sReferNox" & _
               " AND b.sSourceCd = " & strParm(xeSrcMCOrder) & _
               " AND a.sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
               " AND a.nCredtChk < b.nTranAmtx" & _
               " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%")
   lsSQL = lsSQL & _
            " UNION SELECT a.sTransNox" & _
               ", b.sORNoxxxx sReferNox" & _
               ", a.dTransact" & _
               ", b.nTranAmtx - a.nCredtAmt nTranAmtx" & _
               ", a.nCredtAmt" & _
               ", b.cTranStat" & _
               ", IFNULL(c.sTransNox, '') xCheckTrn" & _
               ", IFNULL(c.cDepositd, '1') cDepositd" & _
               ", IFNULL(c.cChckStat, '2') cChckStat" & _
               ", " & strParm(xeSrcOReceipt) & " xSrcTrans" & _
            " FROM MC_Customer_Order a" & _
               ", Receipt_Master b" & _
                  " LEFT JOIN Checks_Received c" & _
                     " ON b.sTransNox = c.sReferNox" & _
                        " AND c.sSourceCd = " & strParm(xeSrcOReceipt) & _
            " WHERE a.sTransNox = b.sReferNox" & _
               " AND b.sSourceCd = " & strParm(xeSrcMCOrder) & _
               " AND a.sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
               " AND a.nCredtAmt < b.nTranAmtx" & _
               " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
            " ORDER BY dTransact" & _
               ", sTransNox"
                
'      lsSQL = "SELECT sTransNox" & _
'                  ", sPRNoxxxx" & _
'                  ", dTransact" & _
'                  ", nTranAmtx" & _
'                  ", cTranStat" & _
'                  ", 'PR' cRcptType" & _
'               " FROM Provisionary_Receipt_Master" & _
'               " WHERE sClientID = " & strParm(.Master("sClientID")) & _
'                  " AND sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
'                  " AND cTranType = " & strParm(0) & _
'                  " AND ( cTranStat = " & strParm(xeStateOpen) & _
'                        " OR (cTranStat IN (" & strParm(xeStateClosed) & ", " & strParm(xeStatePosted) & ")" & _
'                           " AND sReferNox = " & strParm(.Master("sTransNox")) & _
'                           " AND sSourceCd = " & strParm("MCSl") & "))" & _
'               " ORDER BY dTransact"
      
   Set loTmp = New Recordset
   loTmp.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Debug.Print lsSQL
      
   Call createAdvPaym
   With p_oAdvPaym
      Do Until loTmp.EOF
         .AddNew
         
         .Fields("sTransNox") = loTmp("sTransNox")
         .Fields("sReferNox") = loTmp("sReferNox")
         .Fields("dTransact") = loTmp("dTransact")
         .Fields("nTranAmtx") = loTmp("nTranAmtx")
         .Fields("nCredtAmt") = 0
         .Fields("cTranStat") = xeStateOpen
         .Fields("sCheckTrn") = loTmp("xCheckTrn")
         .Fields("cDepositd") = loTmp("cDepositd")
         .Fields("cChckStat") = loTmp("cChckStat")
         .Fields("xSrcTrans") = loTmp("xSrcTrans")
         loTmp.MoveNext
      Loop
      
      If .EOF Then GoTo endWithWarning
      If .RecordCount = 1 Then
         '2008.080.28 Kalyptus
         'Compare the transaction amount with the check amount...
         p_nAdvPaymx = IIf(.Fields("nTranAmtx") > p_oBaseClas.Master("nTranTotl"), p_oBaseClas.Master("nTranTotl"), .Fields("nTranAmtx"))
         
         ' XerSys 2014-04-01
         '  Check if advance payment is paid by check and OR was already released
         
'         If loTmp("xSrcTrans") = xeSrcOReceipt And loTmp("xCheckTrn") <> "" Then
'            p_nAdvForOR = loTmp("nTranAmtx") - IFNull(loTmp("nCredtAmt"), 0)
'         End If
'
         ' XerSys 2014-08-09
         '  If advance payment is greater than transaction amount, use only
         '     how much is being purchase
         If .Fields("xSrcTrans") = xeSrcPReceipt Then
            If p_oAdvPaym("nTranAmtx") > p_oBaseClas.Master("nTranTotl") Then
               p_nAdvForOR = p_oBaseClas.Master("nTranTotl")
            Else
               p_nAdvForOR = p_oAdvPaym("nTranAmtx")
            End If
         End If
         
         p_sAdvRefer = .Fields("sTransNox")
         .Fields("nCredtAmt") = p_nAdvPaymx
         getAdvancePayment = True
         GoTo endProc
      End If
      
      If p_sAdvRefer <> Empty Then
         Do While .EOF = False
            If InStr(1, p_sAdvRefer, .Fields("sTransNox")) > 0 Then
               .Fields("cTranStat") = xeStateClosed
            End If
            
            p_oAdvPaym.MoveNext
         Loop
      End If
      
      Set loAdvPaym = New frmAdvancePayment
      Set loAdvPaym.AppDriver = p_oAppDrivr
      Set loAdvPaym.Source = p_oAdvPaym
      
      loAdvPaym.ClientName = Master("xFullName")
      loAdvPaym.ClientAddress = Master("xAddressx")
      loAdvPaym.TranTotal = Master("nTranTotl")
      loAdvPaym.Show 1
      
      If loAdvPaym.Cancelled Then
         MsgBox "Advance Payment Confirmation was Cancelled!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      p_oAdvPaym.MoveFirst
      p_nAdvPaymx = 0#
      p_nAdvForOR = 0#
      p_sAdvRefer = ""
      Do While p_oAdvPaym.EOF = False
         If p_oAdvPaym("cTranStat") = xeStateClosed Then
            ' XerSys 2014-08-28
            '  Compare the advance payment with the transaction total
'            If p_oBaseClas.Master("nTranTotl") - p_nAdvPaymx > .Fields("nTranAmtx") Then
'               p_nAdvPaymx = p_nAdvPaymx + p_oAdvPaym("nTranAmtx")
'            Else
'               p_nAdvPaymx = p_oBaseClas.Master("nTranTotl") - p_nAdvPaymx
'            End If
            p_nAdvPaymx = p_nAdvPaymx + p_oAdvPaym("nCredtAmt")
            p_sAdvRefer = p_sAdvRefer & "»" & p_oAdvPaym("sTransNox")
         
            ' XerSys 2014-04-01
            '  Check if advance payment is paid by check and OR was already released
            If p_oAdvPaym("xSrcTrans") = xeSrcPReceipt Then
               ' XerSys 2014-08-09
               '  If advance payment is greater than transaction amount, use only
               '     how much is being purchase
               If p_nAdvForOR + p_oAdvPaym("nTranAmtx") > p_oBaseClas.Master("nTranTotl") Then
                  p_nAdvForOR = p_oBaseClas.Master("nTranTotl")
               Else
                  p_nAdvForOR = p_nAdvForOR + p_oAdvPaym("nTranAmtx")
               End If
            End If
         End If
         p_oAdvPaym.MoveNext
      Loop
      
      p_sAdvRefer = Mid(p_sAdvRefer, 2)
   End With
   
   getAdvancePayment = True
   
endProc:
   Exit Function
endWithWarning:
   MsgBox "No Record Found for this Customer!" & _
            vbCrLf & " Can Not Process Advance Payment", vbCritical, "Warning"
   GoTo endProc
errProc:
   ShowError lsProcName
End Function

Private Sub computeTotal()
   Dim lsProcName As String
   Dim lnCtr As Integer
   
   lsProcName = "computeTotal"
   'On Error GoTo errProc
   
   With p_oBaseClas
      .Master("nTranTotl") = 0
      p_oMCPrice.AddtlLoan = 0
      
      For lnCtr = 0 To .ItemCount - 1
         .Master("nTranTotl") = .Master("nTranTotl") + .Detail(lnCtr, "nUnitPrce")
      Next
      
      ' XerSys - 2013-08-19
      '  include other loan amount to the MA calculator
      If .ItemCount > 1 Then
         p_oMCPrice.AddtlLoan = .Master("nTranTotl") - .Detail(0, "nUnitPrce")
      End If
      
      ' after computing total, raise event to display
      RaiseEvent MasterRetrieved(12)
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Private Function SaveMCCustomerTrans(fnEditMode As xeEditMode) As Boolean
   Dim loCltTrans As clsClientTrans
   Dim lsOldProc As String

   lsOldProc = "SaveMCCustomerTrans"
   'On Error GoTo errProc
   
   Set loCltTrans = New clsClientTrans
   With loCltTrans
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      .ClientType = xeClientCustomer
      .SourceSystem = pxeSystemCode
      Call .InitTransaction
      
      ' now save the transaction
      If .Sales(p_oBaseClas.Master("sTransNox"), _
            p_oBaseClas.Master("sClientID"), _
            p_oBaseClas.Master("dTransact"), _
            p_oBaseClas.Master("nTranTotl"), _
            (p_oBaseClas.Master("nAmtPaidx") + p_nAdvPaymx), _
            fnEditMode) = False Then
         GoTo endProc
      End If
   End With
   
   SaveMCCustomerTrans = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "(" & fnEditMode & ")"
   GoTo endProc
End Function

'Private Function SaveMCCustomer()
'   Dim lsSQL As String
'   Dim loRSMaster As Recordset
'   Dim loRSLedger As Recordset
'   Dim bNewRecord As Boolean
'   Dim lsProcName As String
'   Dim lnCodeSize As Integer
'
'   lsProcName = "SaveMCCustomer()"
'   Debug.Print pxeMODULENAME & "." & lsProcName
'   'On Error GoTo errProc
'
'   lnCodeSize = Len(p_oAppDrivr.BranchCode)
'
'   'Handle MC_Customer Master Table
'   lsSQL = " SELECT" & _
'                   "  sClientID" & _
'                   ", sBranchCd" & _
'                   ", sCPerson1" & _
'                   ", sCPPosit1" & _
'                   ", sTelNoxxx" & _
'                   ", sFaxNoxxx" & _
'                   ", sTermIDxx" & _
'                   ", sRemarksx" & _
'                   ", nCredLmit" & _
'                   ", nCredLimt" & _
'                   ", nABalance" & _
'                   ", dCltSince" & _
'                   ", nLedgerNo" & _
'                   ", cRecdStat" & _
'                   ", sModified" & _
'                   ", dModified" & _
'           " FROM MC_Customer" & _
'           " WHERE sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
'             " AND sBranchCd = " & strParm(Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode)))
'   Set loRSMaster = New Recordset
'   loRSMaster.Open lsSQL, p_oAppDrivr, adOpenKeyset, adLockOptimistic, adCmdText
'   Set loRSMaster.ActiveConnection = Nothing
'
'   If loRSMaster.EOF Then
'      bNewRecord = True
'
'      loRSMaster.AddNew
'      loRSMaster("sClientID") = p_oBaseClas.Master("sClientID")
'      loRSMaster("sBranchCd") = Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))
'      loRSMaster("sCPerson1") = p_oBaseClas.Master("xFullName")
'      loRSMaster("nABalance") = 0
'      loRSMaster("dCltSince") = p_oBaseClas.Master("dTransact")
'      loRSMaster("nLedgerNo") = "0"
'      loRSMaster("cRecdStat") = "1"
'   End If
'
'   loRSMaster("nABalance") = loRSMaster("nABalance") + ((p_oBaseClas.Master("nTranTotl") - p_oBaseClas.MasterOriginal("nTranTotl")) - p_oBaseClas.Master("nTranAmtx"))
'   loRSMaster("nLedgerNo") = Format(Val(loRSMaster("nLedgerNo")) + 1, "0000")
'
'   If bNewRecord = True Then
'      lsSQL = ADO2SQL(loRSMaster, _
'                      "MC_Customer", _
'                      , _
'                      p_oAppDrivr.UserID, _
'                      p_oAppDrivr.ServerDate)
'   Else
'      lsSQL = ADO2SQL(loRSMaster, _
'                      "MC_Customer", _
'                      "sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
'                 " AND sBranchCd = " & strParm(Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))), _
'                      p_oAppDrivr.UserID, _
'                      p_oAppDrivr.ServerDate)
'   End If
'
'   If lsSQL <> "" Then
'      p_oAppDrivr.Execute lsSQL, "MC_Customer"
'   End If
'
'   'Handle the ledger
'   lsSQL = "SELECT" & _
'                  ", sClientID" & _
'                  ", sBranchCd" & _
'                  ", dTransact" & _
'                  ", sSystemCd" & _
'                  ", cClientTp" & _
'                  ", nLedgerNo" & _
'                  ", sSourceCd" & _
'                  ", sSourceNo" & _
'                  ", nCreditxx" & _
'                  ", nDebitxxx" & _
'                  ", nABalance" & _
'                  ", dModified" & _
'           " FROM Client_Ledger" & _
'           " WHERE sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
'             " AND sBranchCd = " & strParm(Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))) & _
'             " AND sSourceNo = " & strParm(p_oBaseClas.Master("sTransNox")) & _
'             " AND sSourceCd = " & strParm("MCSl")
'
'   Set loRSLedger = New Recordset
'   loRSLedger.Open lsSQL, p_oAppDrivr, adOpenKeyset, adLockOptimistic, adCmdText
'   Set loRSLedger.ActiveConnection = Nothing
'
'   If loRSLedger.EOF Then
'      bNewRecord = True
'
'      loRSLedger.AddNew
'      loRSLedger("sClientID") = p_oBaseClas.Master("sClientID")
'      loRSLedger("sBranchCd") = Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))
'   End If
'
'   loRSLedger("dTransact") = p_oBaseClas.Master("dTransact")
'   loRSLedger("sSystemCD") = "MC"
'   loRSLedger("sSourceCd") = "MCSl"
'   loRSLedger("sSourceNo") = p_oBaseClas.Master("sTransNox")
'   loRSLedger("nLedgerNo") = loRSMaster("nLedgerNo")
'   loRSLedger("nCreditxx") = 0
'   loRSLedger("nDebitxxx") = p_oBaseClas.Master("nTranTotl")
'   loRSLedger("nABalance") = loRSMaster("nABalance")
'
'   If bNewRecord = True Then
'      lsSQL = ADO2SQL(loRSLedger, _
'                      "Client_Ledger", _
'                      , _
'                      , _
'                      p_oAppDrivr.ServerDate)
'   Else
'      lsSQL = ADO2SQL(loRSLedger, _
'                      "Client_Ledger", _
'                      "sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
'                 " AND sBranchCd = " & strParm(Left(p_oBaseClas.Master("sTransNox"), Len(p_oAppDrivr.BranchCode))) & _
'                 " AND sSourceNo = " & strParm(p_oBaseClas.Master("sTransNox")) & _
'                 " AND sSourceCd = " & strParm("MCSl"), _
'                      , _
'                      p_oAppDrivr.ServerDate)
'   End If
'
'   If lsSQL <> "" Then
'      p_oAppDrivr.Execute lsSQL, "Client_Ledger"
'   End If
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName
'   GoTo endProc
'End Function

Private Function getAdvPaym4Ledger()
   Dim lsOldProc As String
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasRef() As String
   Dim lnCtr As Integer
   
   lsOldProc = "getAdvPaym4Ledger"
   'On Error GoTo errProc
   
   p_oLRMaster.Master("nDownTotl") = (p_oBaseClas.Master("nAmtPaidx") + p_nAdvPaymx) - _
                                       IIf(p_oLRMaster.Master("nAcctTerm") = 0, 0, p_nChattelM) - p_nCommAmtx '- p_oBaseClas.Detail(0, "nInsAmtxx")
   p_oLRMaster.Master("nABalance") = p_oLRMaster.Master("nABalance") - p_oLRMaster.Master("nDownTotl")
   
   lasRef = Split(p_sAdvRefer, "»")
   For lnCtr = 0 To UBound(lasRef)
      lsSQL = ", " & strParm(lasRef(lnCtr))
   Next
   
   lsSQL = "SELECT sORNoxxxx" & _
            " FROM Receipt_Master" & _
            " WHERE sTransNox IN (" & Mid(lsSQL, 2) & " )" & _
            " ORDER BY dTransact DESC"
            
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then GoTo endProc
   
   p_oLRMaster.Master("sORNoxxxx") = loRS("sORNoxxxx")

endProc:
   Exit Function
errProc:
   ShowError lsOldProc
   GoTo endProc
End Function

Private Sub Class_Initialize()
   Set p_oLRMaster = New clsLRMaster
   Set p_oGiveAway = New clsGiveAway
   
   p_cTranStat = xeStateUnknown
End Sub

Private Sub Class_Terminate()
   Set p_oLRMaster = Nothing
   Set p_oBaseClas = Nothing
   Set p_oGiveAway = Nothing
End Sub

Private Sub p_oBaseClas_CancelOthers(Cancel As Boolean)
   If p_sAdvRefer <> "" Then
      Cancel = Not resetAdvance
   End If
   
   'check if commission was already released
   If Cancel = False Then
      If commisionReleased = True Then
         MsgBox "Unable to Cancel/Delete agent commission released transaction.", vbCritical, "Warning"
         Cancel = True
      End If
   End If
End Sub

Private Sub p_oBaseClas_DeleteOthers(Cancel As Boolean)
   Dim lsSQL As String
   
   If p_sAdvRefer <> "" Then
      Cancel = Not resetAdvance
      
      lsSQL = "DELETE FROM MC_SO_Reservation" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
      If p_oAppDrivr.Execute(lsSQL, "MC_SO_Reservation") < 0 Then
         MsgBox "Unable to Remove Sales Reservation Info!" & vbCrLf & _
               "Please Inform SEG/SSG about this Matter!", vbCritical, "Warning"
         Cancel = True
      End If
   End If
   
   If p_sAgentNme <> "" Then
      Cancel = Not delAgent
   End If
   
   'check if commission was already released
   If Cancel = False Then
      If commisionReleased = True Then
         MsgBox "Unable to Cancel/Delete agent commission released transaction.", vbCritical, "Warning"
         Cancel = True
      End If
   End If
End Sub

Private Sub p_oBaseClas_DetailRetrieved(ByVal Index As Integer)
   RaiseEvent DetailRetrieved(Index)
End Sub

Private Sub p_oBaseClas_InitMaster()
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_InitMaster()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc

   With p_oBaseClas
      .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                 p_oAppDrivr.Connection, True, p_sBranchCd)
      .Master("dTransact") = p_oAppDrivr.ServerDate
      .Master("sDRNoxxxx") = GetNextReference(.MasterTable, "sDRNoxxxx", "dTransact", _
                                 "sTransNox", p_sBranchCd, p_oAppDrivr.Connection)
      .Master("sTermName") = ""
      .Master("nTranTotl") = 0#
      .Master("nAmtPaidx") = 0#
      .Master("sORNoxxxx") = ""
      .Master("sApproved") = ""
      .Master("sRemarksx") = ""
      .Master("cPaymForm") = 0
      .Master("cTranType") = 1
      .Master("cCustType") = 0
      .Master("cTranStat") = xeStateOpen
      .Master("sApplicNo") = ""
      .Master("sClientID") = ""
      .Master("sCoCltID1") = ""
      .Master("sCoCltID2") = ""
      .Master("sSalesInv") = ""
      .Master("sTermIDxx") = 0
      .Master("dDueDatex") = p_oAppDrivr.ServerDate
      .Master("sAdvRefer") = xeNo
      .Master("nEntryNox") = 0
      .Master("sTermIDxx") = ""
      .Master("cChngOwnr") = "0"
      .Master("cPromoTyp") = "0"
      
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
      .Detail(0, "sModelIDx") = ""
   End With
   
   With p_axReceipt
      .TransNox = GetNextCode("Receipt_Master", "sTransNox", True, _
                     p_oAppDrivr.Connection, True, p_sBranchCd)
      .ORNoxxxx = GetNextReference("Receipt_Master", "sORNoxxxx", "dTransact", _
                                 "sTransNox", p_sBranchCd, p_oAppDrivr.Connection)
      .TranTotl = 0#
      .CashAmtx = 0#
      .CheckAmt = 0#
      .Remarksx = ""
   End With
   
   Set p_oGift = New GiftCheck
   Set p_oGift.AppDriver = p_oAppDrivr
   p_oGift.InitDetail
   
   p_axReceipt.CashAmtx = 0#
   p_axReceipt.CheckAmt = 0#
   p_axReceipt.ORNoxxxx = ""
   p_axReceipt.Remarksx = ""
   p_axReceipt.TransNox = ""
   p_axReceipt.TranTotl = 0#

   p_xChecks.CheckNox = ""
   p_xChecks.AcctNoxx = ""
   p_xChecks.BankIDxx = ""
   p_xChecks.Amountxx = 0#
   p_xChecks.PRNoxxxx = ""

   p_sFullName = ""
   p_sCoCltNm1 = ""
   p_sCoCltNm2 = ""
   p_sAddressx = ""
   p_sApplicNo = ""
   p_sPRNoxxxx = ""
   p_sAdvRefer = ""
   p_nAdvPaymx = 0#
   
   psModelIDs = ""
   pbHasGAway = False
   
   p_sSQLGvwy = ""
   p_sSQLRebt = ""
   p_sSQLSRPx = ""
   
   ' always create a loan receivable
   Call p_oLRMaster.InitAccount
   ' create new client
   Call p_oClient.InitClient
   Call p_oCoClt1.InitClient
   Call p_oCoClt2.InitClient
   
   Call p_oGCCltx.InitClient
      
   ' create new giveaway
   Call p_oGiveAway.NewTransaction
   'always creat an application
   Call p_oLRApplic.InitTransaction
   
   ' Init Agent records
   Call initAgent
   Call initAgentInfo
   Call initAgentTrans
   
   Set p_oClientInfo = New ClientInfo
   Set p_oCoMkrInfo1 = New ClientInfo
   Set p_oCoMkrInfo2 = New ClientInfo
   
   Set p_oGCardAppl = Nothing
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName
End Sub

Private Sub p_oBaseClas_LoadOthers()
   Dim loOReceipt As clsORReceipt
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsClientID As String
   
   lsProcName = "p_oBaseClas_LoadOthers()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   pbChecksOK = True
   With p_oBaseClas
      p_sFullName = .Master("xFullName")
      p_sCoCltNm1 = IFNull(.Master("xCoCltNm1"))
      p_sCoCltNm2 = IFNull(.Master("xCoCltNm2"))
      
      p_sAddressx = .Master("xAddressx")
       
      If .Master("cTranStat") = xeStateCancelled Then GoTo endProc
             
      'Transfer loading of giveaway here...
      pbHotItemx = False
      For pnCtr = 0 To .ItemCount() - 1
         If .Detail(pnCtr, "cMotorNew") = xeYes Then
            psModelIDs = "»" & p_oBaseClas.Detail(pnCtr, "sModelIDx")
         End If
         
         If Not pbHotItemx Then
            If .Detail(pnCtr, "cMotorNew") = xeYes Then pbHotItemx = IIf(.Detail(pnCtr, "cHotItemx") = 0, False, True)
         End If
      Next
      psModelIDs = Mid(psModelIDs, 2)
      
      If p_oGiveAway.OpenTransaction(.Master("sTransNox")) = False Then
         Call p_oGiveAway.NewTransaction
      End If
      
      pbHasGAway = p_oGiveAway.Detail(0, "sPartsIDx") <> Empty
      'Up to this area...
      
      Set loOReceipt = New clsORReceipt
      Set loOReceipt.AppDriver = p_oAppDrivr
      loOReceipt.Branch = p_sBranchCd
      loOReceipt.Parent = "clsMCSales"
            
      lsClientID = p_oBaseClas.Master("sClientID")
      loOReceipt.Filter = "a.sClientID = " & strParm(lsClientID)
      
      If loOReceipt.InitTransaction = False Then GoTo endProc
      If .Master("nAmtPaidx") > 0 Then
         If loOReceipt.SearchByRef(p_oBaseClas.Master("sTransNox"), "MCSl") = False Then
            MsgBox "Unable to Load Receipt Info!", vbCritical, "Warning"
            GoTo endProc
         End If
            
         p_axReceipt.ORNoxxxx = IFNull(loOReceipt.Master("sORNoxxxx"))
         p_axReceipt.CashAmtx = IIf(p_axReceipt.ORNoxxxx = "", 0, loOReceipt.Master("nTranAmtx"))
         p_axReceipt.CheckAmt = loOReceipt.Checks("nAmountxx")
         p_axReceipt.Remarksx = IFNull(loOReceipt.Master("sRemarksx"))
         p_axReceipt.TransNox = IFNull(loOReceipt.Master("sTransNox"))
         p_axReceipt.TranTotl = .Master("nTranTotl")
         p_sPRNoxxxx = loOReceipt.PRNumber
         p_xChecks.CheckNox = loOReceipt.Checks("sCheckNox")
         p_xChecks.AcctNoxx = loOReceipt.Checks("sAcctNoxx")
         p_xChecks.BankIDxx = loOReceipt.Checks("sbankidxx")
         p_xChecks.CheckDte = loOReceipt.Checks("dCheckDte")
         p_xChecks.Amountxx = loOReceipt.Checks("nAmountxx")
         
         If loOReceipt.Checks("cDepositd") = xeYes Then pbChecksOK = False
      End If
      
      'Kalyptus - 2015.06.19 09:07pm
      'Load Gift Check
      Set p_oGift = New GiftCheck
      Set p_oGift.AppDriver = p_oAppDrivr
      p_oGift.InitDetail
      Set p_oGift = loOReceipt.GiftCoupon
      
      Call p_oLRMaster.InitAccount
      
      p_sApplicNo = .Master("sApplicNo")
      If p_sApplicNo = "N-O-N-E" Then
         Call p_oLRMaster.LoadByCondition( _
               "a.sClientID = " & strParm(.Master("sClientID")) & _
               " AND a.sSerialID = " & strParm(.Detail(0, "sSerialID")))
      ElseIf p_sApplicNo <> "" Then
         Call p_oLRMaster.LoadByCondition("sApplicNo = " & strParm(.Master("sApplicNo")))
      End If
      
      p_oClient.InitClient
      p_oCoClt1.InitClient
      p_oCoClt2.InitClient
      p_oGCCltx.InitClient
      
      p_sAdvRefer = ""
      p_nAdvPaymx = 0
      
      ' check if advance payment exist
      If .Master("sAdvRefer") = CStr(xeYes) Then
         Call createAdvPaym
         
         lsSQL = "SELECT a.sTransNox" & _
                     ", b.sPRNoxxxx sReferNox" & _
                     ", a.dTransact" & _
                     ", d.nAmountxx" & _
                     ", b.cTranStat" & _
                     ", IFNULL(c.sTransNox, '') xCheckTrn" & _
                     ", IFNULL(c.cDepositd, '1') cDepositd" & _
                     ", IFNULL(c.cChckStat, '2') cChckStat" & _
                     ", " & strParm(xeSrcPReceipt) & " xSrcTrans" & _
                  " FROM MC_Customer_Order a" & _
                     ", Provisionary_Receipt_Master b" & _
                        " LEFT JOIN Checks_Received c" & _
                           " ON b.sTransNox = c.sReferNox" & _
                              " AND c.sSourceCd = " & strParm(xeSrcPReceipt) & _
                     ", MC_SO_Reservation d" & _
                  " WHERE a.sTransNox = b.sReferNox" & _
                     " AND b.sSourceCd = " & strParm(xeSrcMCOrder) & _
                     " AND a.sTransNox = d.sReferNox" & _
                     " AND d.sSourceCd = " & strParm(xeSrcPReceipt) & _
                     " AND d.sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
         lsSQL = lsSQL & _
                  " UNION SELECT a.sTransNox" & _
                     ", b.sORNoxxxx sReferNox" & _
                     ", a.dTransact" & _
                     ", d.nAmountxx" & _
                     ", b.cTranStat" & _
                     ", IFNULL(c.sTransNox, '') xCheckTrn" & _
                     ", IFNULL(c.cDepositd, '1') cDepositd" & _
                     ", IFNULL(c.cChckStat, '2') cChckStat" & _
                     ", " & strParm(xeSrcOReceipt) & " xSrcTrans" & _
                  " FROM MC_Customer_Order a" & _
                     ", Receipt_Master b" & _
                        " LEFT JOIN Checks_Received c" & _
                           " ON b.sTransNox = c.sReferNox" & _
                              " AND c.sSourceCd = " & strParm(xeSrcOReceipt) & _
                     ", MC_SO_Reservation d" & _
                  " WHERE a.sTransNox = b.sReferNox" & _
                     " AND b.sSourceCd = " & strParm(xeSrcMCOrder) & _
                     " AND a.sClientID = " & strParm(.Master("sClientID")) & _
                     " AND a.sTransNox = d.sReferNox" & _
                     " AND d.sSourceCd = " & strParm(xeSrcOReceipt) & _
                     " AND d.sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                  " ORDER BY dTransact" & _
                     ", sTransNox"
         Set loRS = New Recordset
         loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
         
         With p_oAdvPaym
            Do Until loRS.EOF
               .AddNew
               
               .Fields("sTransNox") = loRS("sTransNox")
               .Fields("sReferNox") = loRS("sReferNox")
               .Fields("dTransact") = loRS("dTransact")
               .Fields("nTranAmtx") = loRS("nAmountxx")
               .Fields("nCredtAmt") = loRS("nAmountxx")
               .Fields("cTranStat") = xeStateClosed
               .Fields("sCheckTrn") = loRS("xCheckTrn")
               .Fields("cDepositd") = loRS("cDepositd")
               .Fields("cChckStat") = loRS("cChckStat")
               .Fields("xSrcTrans") = loRS("xSrcTrans")
               
               p_nAdvPaymx = p_nAdvPaymx + loRS("nAmountxx")
               p_sAdvRefer = "»" & .Fields("sTransNox")
               loRS.MoveNext
            Loop
            Debug.Print lsSQL
            p_sAdvRefer = Mid(p_sAdvRefer, 2)
         End With
         
         ' XerSys 2014-03-26
         '  Disable the following codes
'         lsSQL = "SELECT sTransNox" & _
'                     ", sPRNoxxxx" & _
'                     ", dTransact" & _
'                     ", nTranAmtx" & _
'                     ", cTranStat" & _
'                     ", 'PR' cRcptType" & _
'                  " FROM Provisionary_Receipt_Master" & _
'                  " WHERE sReferNox = " & strParm(.Master("sTransNox")) & _
'                     " AND sSourceCd = " & strParm("MCSl") & _
'                  " UNION SELECT sTransNox" & _
'                     ", sORNoxxxx sPRNoxxxx" & _
'                     ", dTransact" & _
'                     ", nTranAmtx" & _
'                     ", cTranStat" & _
'                     ", 'OR' cRcptType" & _
'                  " FROM Receipt_Master" & _
'                  " WHERE cTranType = " & strParm(0) & _
'                     " AND sReferNox = " & strParm(.Master("sTransNox")) & _
'                     " AND sSourceCd = " & strParm("MCSl") & _
'                  " ORDER BY dTransact"
'         Set loRS = New Recordset
'         loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
'         p_nAdvPaymx = 0#
'         p_sAdvRefB4 = ""
'
'         Do While loRS.EOF = False
'            p_nAdvPaymx = p_nAdvPaymx + loRS("nTranAmtx")
'            p_sAdvRefB4 = p_sAdvRefB4 & "»" & loRS("sTransNox") & loRS("cRcptType")
'
'            loRS.MoveNext
'         Loop
'         p_sAdvRefB4 = Mid(p_sAdvRefB4, 2)
      End If
      
      ' Init Agent records
      Call initAgent
      Call initAgentInfo
      Call initAgentTrans
      Call loadAgentTrans
      Call loadLTMSInfo
      
   End With
endProc:
   Set loOReceipt = Nothing
   Set loRS = Nothing
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub

Private Sub loadLTMSInfo()
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsSQL = "SELECT" & _
               "  a.sLastName" & _
               ", a.sFrstName" & _
               ", a.sMiddName" & _
               ", a.sSuffixNm" & _
               ", a.dBirthDte" & _
               ", IFNULL(b.sTownName, '') sBrthTown" & _
               ", a.sHouseNox" & _
               ", a.sAddressx" & _
               ", IFNULL(c.sBrgyName, '') sBrgyName" & _
               ", IFNULL(d.sTownName, '') xTownName" & _
               ", IFNULL(e.sProvName, '') xProvName" & _
               ", TRIM(CONCAT(a.sHouseNox, ' ', a.sAddressx, ' ', IFNULL(d.`sTownName`, ''), ', ', IFNULL(e.`sProvName`, ''), ' ', IFNULL(d.`sZippCode`, ''))) xAddressx" & _
               ", a.sMobileNo" & _
               ", a.sEmailAdd" & _
               ", IFNULL(f.sLTOIDxxx, 'NONE') sLTOIDxxx" & _
               ", IFNULL(a.sFatherID, '') sFatherID" & _
               ", IFNULL(a.sMotherID, '') sMotherID" & _
               ", IFNULL(g.sOccptnNm, '') sOccptnNm" & _
               ", IFNULL(h.sNational, '') sCitizenx" & _
               ", a.cGenderCd" & _
               ", a.sTaxIDNox"
   lsSQL = lsSQL & _
            " FROM Client_Master a" & _
               " LEFT JOIN TownCity b ON a.sBirthPlc = b.sTownIDxx" & _
               " LEFT JOIN Barangay c ON a.sBrgyIDxx = c.sBrgyIDxx" & _
               " LEFT JOIN TownCity d ON a.sTownIDxx = d.sTownIDxx" & _
               " LEFT JOIN Province e ON d.sProvIDxx = e.sProvIDxx" & _
               " LEFT JOIN Client_Others f ON a.sClientID = f.sClientID" & _
               " LEFT JOIN Occupation g ON a.sOccptnID = g.sOccptnID" & _
               " LEFT JOIN Country h ON a.sCitizenx = h.sCntryCde"

   Set loRS = New Recordset
   loRS.Open AddCondition(lsSQL, "a.sClientID = " & strParm(Master("sClientID"))), p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing

   Set p_oClientInfo = New ClientInfo
   p_oClientInfo.sLastName = loRS("sLastName")
   p_oClientInfo.sFrstName = loRS("sFrstName")
   p_oClientInfo.sMiddName = loRS("sMiddName")
   p_oClientInfo.sSuffixNm = loRS("sSuffixNm")
   p_oClientInfo.sBirthDte = loRS("dBirthDte")
   p_oClientInfo.sBirthPlc = loRS("sBrthTown")
   p_oClientInfo.sHouseNox = loRS("sHouseNox")
   p_oClientInfo.sAddresxx = loRS("sAddressx")
   p_oClientInfo.sBarangay = loRS("sBrgyName")
   p_oClientInfo.sTownxxxx = loRS("xTownName")
   p_oClientInfo.sProvince = loRS("xProvName")
   p_oClientInfo.xAddressx = loRS("xAddressx")
   p_oClientInfo.sMobileNo = loRS("sMobileNo")
   p_oClientInfo.sEmailAdd = loRS("sEmailAdd")
   p_oClientInfo.sLTOIDxxx = loRS("sLTOIDxxx")
   p_oClientInfo.sFatherNm = loRS("sFatherID")
   p_oClientInfo.sMotherNm = loRS("sMotherID")
   p_oClientInfo.sOccptnNm = loRS("sOccptnNm")
   p_oClientInfo.sCitizenx = loRS("sCitizenx")
   p_oClientInfo.sGenderCd = loRS("cGenderCd")
   p_oClientInfo.sTaxIDNox = loRS("sTaxIDNox")
   
   Set p_oCoMkrInfo1 = New ClientInfo
   If IFNull(Master("sCoCltID1")) <> "" Then
      Set loRS = New Recordset
      loRS.Open AddCondition(lsSQL, "a.sClientID = " & strParm(Master("sCoCltID1"))), p_oAppDrivr.Connection, , , adCmdText
      Set loRS.ActiveConnection = Nothing
   
      Set p_oCoMkrInfo1 = New ClientInfo
      p_oCoMkrInfo1.sLastName = loRS("sLastName")
      p_oCoMkrInfo1.sFrstName = loRS("sFrstName")
      p_oCoMkrInfo1.sMiddName = loRS("sMiddName")
      p_oCoMkrInfo1.sSuffixNm = loRS("sSuffixNm")
      p_oCoMkrInfo1.sBirthDte = loRS("dBirthDte")
      p_oCoMkrInfo1.sBirthPlc = loRS("sBrthTown")
      p_oCoMkrInfo1.sHouseNox = loRS("sHouseNox")
      p_oCoMkrInfo1.sAddresxx = loRS("sAddressx")
      p_oCoMkrInfo1.sBarangay = loRS("sBrgyName")
      p_oCoMkrInfo1.sTownxxxx = loRS("xTownName")
      p_oCoMkrInfo1.sProvince = loRS("xProvName")
      p_oCoMkrInfo1.xAddressx = loRS("xAddressx")
      p_oCoMkrInfo1.sMobileNo = loRS("sMobileNo")
      p_oCoMkrInfo1.sEmailAdd = loRS("sEmailAdd")
      p_oCoMkrInfo1.sLTOIDxxx = loRS("sLTOIDxxx")
      p_oCoMkrInfo1.sFatherNm = loRS("sFatherID")
      p_oCoMkrInfo1.sMotherNm = loRS("sMotherID")
      p_oCoMkrInfo1.sOccptnNm = loRS("sOccptnNm")
      p_oCoMkrInfo1.sCitizenx = loRS("sCitizenx")
      p_oCoMkrInfo1.sGenderCd = loRS("cGenderCd")
      p_oCoMkrInfo1.sTaxIDNox = loRS("sTaxIDNox")
   End If
   
   Set p_oCoMkrInfo2 = New ClientInfo
   If IFNull(Master("sCoCltID2")) <> "" Then
      Set loRS = New Recordset
      loRS.Open AddCondition(lsSQL, "a.sClientID = " & strParm(Master("sCoCltID2"))), p_oAppDrivr.Connection, , , adCmdText
      Set loRS.ActiveConnection = Nothing
   
      Set p_oCoMkrInfo2 = New ClientInfo
      p_oCoMkrInfo2.sLastName = loRS("sLastName")
      p_oCoMkrInfo2.sFrstName = loRS("sFrstName")
      p_oCoMkrInfo2.sMiddName = loRS("sMiddName")
      p_oCoMkrInfo2.sSuffixNm = loRS("sSuffixNm")
      p_oCoMkrInfo2.sBirthDte = loRS("dBirthDte")
      p_oCoMkrInfo2.sBirthPlc = loRS("sBrthTown")
      p_oCoMkrInfo2.sHouseNox = loRS("sHouseNox")
      p_oCoMkrInfo2.sAddresxx = loRS("sAddressx")
      p_oCoMkrInfo2.sBarangay = loRS("sBrgyName")
      p_oCoMkrInfo2.sTownxxxx = loRS("xTownName")
      p_oCoMkrInfo2.sProvince = loRS("xProvName")
      p_oCoMkrInfo2.xAddressx = loRS("xAddressx")
      p_oCoMkrInfo2.sMobileNo = loRS("sMobileNo")
      p_oCoMkrInfo2.sEmailAdd = loRS("sEmailAdd")
      p_oCoMkrInfo2.sLTOIDxxx = loRS("sLTOIDxxx")
      p_oCoMkrInfo2.sFatherNm = loRS("sFatherID")
      p_oCoMkrInfo2.sMotherNm = loRS("sMotherID")
      p_oCoMkrInfo2.sOccptnNm = loRS("sOccptnNm")
      p_oCoMkrInfo2.sCitizenx = loRS("sCitizenx")
      p_oCoMkrInfo2.sGenderCd = loRS("cGenderCd")
      p_oCoMkrInfo2.sTaxIDNox = loRS("sTaxIDNox")
   End If
End Sub

Private Sub p_oBaseClas_MasterRetrieved(ByVal Index As Integer)
   RaiseEvent MasterRetrieved(Index)
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = Not pbModified
End Sub

Private Sub p_oBaseClas_SaveComplete()
   Dim lsProcName As String
   Dim loJson As New rmjJSON.clsJSON
   Dim loObjx As Object
   Dim lsResult As String
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsProcName = "p_oBaseClas_SaveComplete()"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If p_oGCardAppl Is Nothing Then
      GoTo endProc
   End If

   If Not (p_oGCardAppl("cDigitalx") = "1") Then
      GoTo endProc
   End If
   
   'check for mobile number
   If Trim(p_oClient.Master("sMobileNo")) = "" Then
      MsgBox "Empty mobile number detected for this client." & vbCrLf & _
         "Please update his mobile number to continue registering on DIGITAL CARD." & vbCrLf & vbCrLf & _
         "The customer will receive a text message if his application was approved."
      GoTo endProc
   End If
   
   'mac 2020.08.12
   '  get server address on the database config
   lsSQL = p_oAppDrivr.getConfiguration("WebSvr")
   
   'check if the branch is connected to the main server
   If Not RMJConnected(lsSQL) Then
      MsgBox "You are not connected to the main server..." & vbCrLf & _
               "Please inform the customer to wait for our text regarding his GCARD INFORMATION...", vbInformation, "Notice"
      GoTo endProc
   End If
   
   Set loRS = New Recordset
   With loRS
      .Fields.Append "transnox", adVarChar, 100
      .Fields.Append "transact", adVarChar, 100
      .Fields.Append "sourcecd", adVarChar, 100
      .Fields.Append "purcmode", adVarChar, 100
      .Fields.Append "clientid", adVarChar, 100
      .Fields.Append "cardtype", adVarChar, 100
      .Fields.Append "appltype", adVarChar, 100
      .Fields.Append "prevgcrd", adVarChar, 100
      .Fields.Append "sourceno", adVarChar, 100
      .Fields.Append "amtpaidx", adInteger
      .Fields.Append "compnyid", adVarChar, 100
      .Fields.Append "serialid", adVarChar, 100
      .Fields.Append "yellowxx", adInteger
      .Fields.Append "whitexxx", adInteger
      .Fields.Append "pointsxx", adInteger
      .Fields.Append "digitalx", adVarChar, 100
      .Open
   
      .AddNew
      .Fields("transnox") = p_oGCardAppl("sTransNox")
      .Fields("transact") = Format(p_oGCardAppl("dTransact"), "YYYY-MM-DD")
      .Fields("sourcecd") = p_oGCardAppl("sSourceCd")
      .Fields("purcmode") = p_oGCardAppl("cPurcMode")
      .Fields("clientid") = p_oGCardAppl("sClientID")
      .Fields("cardtype") = p_oGCardAppl("cCardType")
      .Fields("appltype") = p_oGCardAppl("cApplType")
      .Fields("prevgcrd") = ""
      .Fields("sourceno") = p_oGCardAppl("sSourceNo")
      .Fields("amtpaidx") = 0
      .Fields("compnyid") = p_oGCardAppl("sCompnyID")
      .Fields("serialid") = p_oGCardAppl("sSerialID")
      .Fields("yellowxx") = p_oGCardAppl("nYellowxx")
      .Fields("whitexxx") = p_oGCardAppl("nWhitexxx")
      .Fields("pointsxx") = 0
      .Fields("digitalx") = "1"
   End With
         
   Set loObjx = loJson.parse(ggcRestAPI.ApproveApplication(p_oAppDrivr, loRS, p_oClient.Master("sMobileNo")))
      
   If (loObjx.Item("result") = "SUCCESS") Then
      lsSQL = "UPDATE G_Card_Application" & _
             " SET cTranStat = '1'" & _
             " WHERE sTransNox = " & strParm(p_oGCardAppl("sTransNox"))
      p_oAppDrivr.Execute lsSQL, "G_Card_Application"
      
      lsResult = loObjx.Item("sCardNmbr") 'get the gcard number value
      Debug.Print (lsResult)
      
      lsResult = "APPLICATION»»»»" + lsResult + "»»»»»»" 'insert to format
      Debug.Print (lsResult)
      
      lsResult = StringToHex(lsResult) 'convert to hexadecimal
      Debug.Print (lsResult)
      
      newcard lsResult 'show QR CODE
   Else
      Set loObjx = loObjx.Item("error")
      MsgBox loObjx.Item("message")
   End If
      
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Private Sub p_oBaseClas_SaveOthers(Cancel As Boolean)
   Dim loCltTrans As clsClientTrans
   Dim loFinTrans As clsFinanceMaster
   Dim lsProcName As String
   Dim lbCancel As Boolean
   Dim lsSQL As String
   
   
   lsProcName = "p_oBaseClas_SaveOthers(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True
   
   With p_oBaseClas
      'mac 2021.02.26
      '  removed this block of code since we are not using the MAPFRE insurance now.
'      'iMac [2015-05-20]
'      'MAPFRE Insurance here:
'      If .EditMode = xeModeAddNew Then
'         If IFNull(p_oAppDrivr.getConfiguration("cMAPFREIn", p_oAppDrivr.BranchCode), "0") = xeYes Then
'            If Trim(.Detail(0, "sSerialID")) <> "" And .Detail(0, "cMotorNew") = xeYes Then
'               If saveAssuredInfo() = False Then GoTo endProc
'            End If
'         End If
'      End If

      'Save giveaway approval here....
      If p_sSQLGvwy <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
      
         p_oAppDrivr.Execute Replace(p_sSQLGvwy, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
   
      'Save rebate approval here...
      If p_sSQLRebt <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
         p_oAppDrivr.Execute Replace(p_sSQLRebt, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
   
      'Save MC Price Discount...
      If p_sSQLSRPx <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
         p_oAppDrivr.Execute Replace(p_sSQLSRPx, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
      
      'mac 2019.10.18
      '  save first pay approval code
      If p_sSQLFPay <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
         p_oAppDrivr.Execute Replace(p_sSQLFPay, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
      
      'mac 2020.07.03
      '  save mc sales w/o pnp csr approval code
      If p_sSQLCSRx <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
         p_oAppDrivr.Execute Replace(p_sSQLCSRx, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
      
      'mac 2020.07.31
      '  update status of Credit Online Application to USED
      If p_sSQgocas <> "" Then
         p_oAppDrivr.Execute p_sSQgocas, "Credit_Online_Application"
      End If
      
      'mac 2020.08.24
      '  save mc sales commission approval code
      If p_sSQLComm <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
         p_oAppDrivr.Execute Replace(p_sSQLComm, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
      
      'mac 2021.08.24
      '  save advance/back date approval code
      If p_sSQLDate <> "" Then
         lsSQL = GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                          p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      
         p_oAppDrivr.Execute Replace(p_sSQLDate, "xTransNox", lsSQL), "xxxSCA_Usage"
      End If
      
      ' save motorcycle
      If saveMCTransaction() = False Then GoTo endProc

      ' then save receipt info
      'If .Master("nAmtPaidx") + p_nAdvPaymx > 0 Then
      If .Master("nAmtPaidx") + p_nAdvForOR > 0 Then
         'If saveReceipt(xeStateOpen, .Master("nAmtPaidx") + p_nAdvPaymx) = False Then GoTo endProc
         If saveReceipt(xeStateOpen, .Master("nAmtPaidx") + p_nAdvForOR) = False Then GoTo endProc
      ElseIf p_nAdvPaymx > 0 Then
         ' XerSys - 2015-12-21
         '  Manually retrieve the or number of advance payment for the account ledger
         If p_oBaseClas.Master("sApplicNo") <> Empty And _
            p_oBaseClas.Master("cPaymForm") <> "0" Then
            Call getAdvPaym4Ledger
         End If
      ElseIf .Master("cPaymForm") = "4" Then    ' Term with Registration
         If saveReceipt(xeStateOpen, .Master("nTranTotl")) = False Then GoTo endProc
      End If

      ' XerSys 2014-03-26
      '  Remove the p_sAdvRefB4 in the condtionion
      'If p_sAdvRefer <> Empty Or p_sAdvRefB4 <> Empty Then
      If p_sAdvRefer <> Empty Then
         If updateAdvPayment() = False Then GoTo endProc
      End If

      ' now save account if mc was sold as installment or cash balance
      If .Master("sApplicNo") <> Empty And _
         .Master("cPaymForm") <> "0" Then
         ' Assign Application number to L/R
         p_oLRMaster.Master("sApplicNo") = .Master("sApplicNo")
         p_oLRMaster.Master("sClientID") = .Master("sClientID")
         p_oLRMaster.Master("sCoCltID1") = .Master("sCoCltID1")
         p_oLRMaster.Master("sCoCltID2") = .Master("sCoCltID2")
         
         If p_oLRMaster.SaveAccount() = False Then GoTo endProc
         
      ElseIf p_sApplicNo <> Empty Then
         ' @@@ Delete Application
         If p_oLRApplic.DeleteTransaction = False Then GoTo endProc
         If p_oLRMaster.DeleteAccount() = False Then GoTo endProc
      End If

      If .EditMode = xeModeUpdate Then
         'Previously entered as Term/Term with Registration
         If (.MasterOriginal("cPaymForm") = "3" Or .MasterOriginal("cPaymForm") = "4") _
              And Not (.Master("cPaymForm") = "3" Or .Master("cPaymForm") = "4") Then
            
            'Delete the ledger
            If SaveMCCustomerTrans(xeModeDelete) = False Then GoTo endProc
            
         'Previously entered as finance
         ElseIf .MasterOriginal("cPaymForm") = "5" And Not .Master("cPaymForm") = 5 Then
            'Delete info of this financing account
            Set loFinTrans = New clsFinanceMaster
            Set loFinTrans.AppDriver = p_oAppDrivr
            loFinTrans.Branch = p_sBranchCd
            loFinTrans.Parent = pxeMODULENAME
            Call loFinTrans.InitAccount
         
            If loFinTrans.LoadByCondition("sReferNox = " & strParm(p_oBaseClas.Master("sTransNox"))) = True Then
               loFinTrans.DeleteAccount
            End If
         End If
      End If

      ' only institutional sales has its transaction ledger
'      If .Master("cCustType") = 1 Then
      'All term payment form should have a ledger
      If .Master("cPaymForm") = "3" Or .Master("cPaymForm") = "4" Then
         'Save ledger
         If SaveMCCustomerTrans(.EditMode) = False Then GoTo endProc
      ElseIf .Master("cPaymForm") = "5" Then
         'Save info of this financing account
         'Assume one serial for each dr...
         Set loFinTrans = New clsFinanceMaster
         Set loFinTrans.AppDriver = p_oAppDrivr
         loFinTrans.Branch = p_sBranchCd
         loFinTrans.Parent = pxeMODULENAME
         Call loFinTrans.InitAccount
      
         If p_oBaseClas.EditMode <> xeModeAddNew Then
            If loFinTrans.LoadByCondition("sReferNox = " & strParm(.Master("sTransNox"))) = False Then
               loFinTrans.NewAccount
            End If
         Else
            loFinTrans.NewAccount
         End If
                  
         loFinTrans.Master("sReferNox") = .Master("sTransNox")
         loFinTrans.Master("sDRNoxxxx") = .Master("sDRNoxxxx")
         loFinTrans.Master("dTransact") = .Master("dTransact")
         loFinTrans.Master("sClientID") = .Master("sClientID")
         loFinTrans.Master("sCoCltID1") = .Master("sCoCltID1")
         loFinTrans.Master("sCoCltID2") = .Master("sCoCltID2")
         loFinTrans.Master("sSerialID") = .Detail(0, "sSerialID")
         loFinTrans.Master("nCashPrce") = .Master("nTranTotl")
         loFinTrans.Master("nDownPayM") = .Master("nAmtPaidx")
         
         loFinTrans.SaveAccount
      End If
   End With
   
   'Kalyptus - 2015.12.22 11:12am
   'Save GCard Here!!!
   If Not (p_oGCardAppl Is Nothing) Then
      
      'Kalyptus - 2016.06.17 05:03pm
      'Check if mc client and gcard client are the same and set the Client ID accordingly
      If p_oGCCltx Is Nothing Then
         p_oGCardAppl("sClientID") = p_oBaseClas.Master("sClientID")
      Else
         If p_oGCCltx.Master("sLastName") <> "" Then
            If Not p_oGCCltx.SaveClient Then GoTo endProc
            p_oGCardAppl("sClientID") = p_oGCCltx.Master("sClientID")
         Else
            p_oGCardAppl("sClientID") = p_oBaseClas.Master("sClientID")
         End If
      End If
      
      p_oGCardAppl("sTransNox") = GetNextCode("G_Card_Application", "sTransNox", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
      lsSQL = ADO2SQL(p_oGCardAppl, "G_Card_Application", , p_oAppDrivr.UserID, p_oAppDrivr.ServerDate)
      Call p_oAppDrivr.Execute(lsSQL, "G_Card_Application")
   End If

   ' process giveaway after saving other info
   With p_oGiveAway
      .Master("sTransNox") = p_oBaseClas.Master("sTransNox")
      If .Detail(0, "sPartsIDx") <> Empty Then
         .Master("sClientID") = p_oBaseClas.Master("sClientID")
         .Master("dTransact") = p_oBaseClas.Master("dTransact")
         If .SaveTransaction() = False Then GoTo endProc
      ElseIf pbHasGAway Then
         If .DeleteTransaction() = False Then GoTo endProc
      End If
      
      If p_sAgentNme <> "" Then
         If Not updateAgent() Then GoTo endProc
      End If
   End With
   
   Cancel = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub

Private Sub p_oBaseClas_WillCancel(Cancel As Boolean)
   Dim lsProcName As String
   
   lsProcName = "p_oBaseClas_WillCancel(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
            
   Cancel = Not delOtherTrans()

endProc:
   Exit Sub
errProc:
    ShowError lsProcName
End Sub

Private Sub p_oBaseClas_WillDelete(Cancel As Boolean)
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsProcName = "p_oBaseClas_WillDelete(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   Cancel = Not delOtherTrans()
   
endProc:
   Exit Sub
errProc:
    ShowError lsProcName
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim lsProcName As String
   Dim lbCancel As Boolean
   Dim lsDetail As String
   Dim lsApprovedCD As String
   Dim lsApproveID As String
   Dim lsApproveName As String

   lsProcName = "p_oBaseClas_WillSave(" & Cancel & ")"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   Cancel = True

   pbModified = True
   With p_oBaseClas
      ' check if transaction detail exist
      If .Detail(0, "sEngineNo") = "" And .Detail(0, "sRemarksx") = Empty Then
         MsgBox "Invalid Transaction Detail Detected!!!" & vbCrLf & _
                  "Verify your Entry then Try Again!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      'mac 2020.08.26
      If Not validSalesCommision Then GoTo endProc
      
      ' XerSys - 2016-09-23
      '  validate SRP of 2H
      If .Detail(0, "sSerialID") <> "" And .Detail(0, "cMotorNew") = xeNo Then
         If Not valid2HPrice() Then GoTo endProc
      Else
         'kalyptus - 2016.02.13 10:31am
         Select Case .Master("cPaymForm")
         Case "0"    ' Cash
            If Not (.Detail(0, "sSerialID") = "" Or .Detail(0, "cMotorNew") = xeNo) Then
               If Not checkSRP(.Master("cPaymForm")) Then GoTo endProc
            End If
         Case "1", "2"  ' Installment; Cash Balance
            If Not (.Detail(0, "sSerialID") = "" Or .Detail(0, "cMotorNew") = "") Then
               If Not checkSRP(.Master("cPaymForm")) Then GoTo endProc
            End If
         End Select
      End If
      
      'mac 2021.09.21
      'validate the date if back/advance date
      p_sSQLDate = ""
      pnCtr = validateTranDate(p_oAppDrivr, .Master("dTransact"))
      If pnCtr <> 0 Then
         Select Case pnCtr
         Case 1
            MsgBox "You will be asked for an ADVANCE DATE approval code." & vbCrLf & vbCrLf & _
                  "Please ask approval to CSS Department.", vbInformation, "Notice"
         Case 2
            MsgBox "You will be asked for an BACK DATE approval code." & vbCrLf & vbCrLf & _
                  "Please ask approval to MIS Department.", vbInformation, "Notice"
         Case Else
            MsgBox "Invalid transaction date detected.", vbCritical, "Warning"
            GoTo endProc
         End Select
         
         If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
            MsgBox "User is not authorized to Approve Advance Date of MC Sales!" & vbCrLf & _
                  "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
            GoTo endProc
         End If
                  
         If isValidApproveCode( _
               p_oSCA_Const.pxeAdvanceBackDate, _
               p_oAppDrivr.BranchCode, _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               Master("sDRNoxxxx"), _
               lsApprovedCD) Then
               
            p_sSQLDate = "xTransNox"
            p_sSQLDate = "INSERT INTO xxxSCA_Usage" & _
                        " SET sTransNox = " & strParm(p_sSQLDate) & _
                           ", sApprCode = " & strParm(lsApprovedCD) & _
                           ", sApproved = " & strParm(lsApproveID) & _
                           ", sSystemCD = " & strParm(p_oSCA_Const.pxeAdvanceBackDate) & _
                           ", sSourceNo = " & strParm(Master("sTransNox")) & _
                           ", sSourceCD = " & strParm(pxeMCSales) & _
                           ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                           ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
         Else
            MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
               "Verify entry then try again!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
      
      p_sSQLCSRx = ""
      'kalyptus - 2020.06.29 12:03pm
      'Check for the cCSRValid of the Item Sold
      If Val(IFNull(.Detail(0, "cCSRValid"), 0)) = 0 Then
         lsDetail = p_oAppDrivr.getConfiguration("SOWoPNP")
         
         'Is sales of MCs without CSR not allowed
         If Val(lsDetail) = 0 Then
            MsgBox "Selling of MC(s) without CSR is not allowed!", vbCritical, "Warning"
            GoTo endProc
         'Is sales of MCs without CSR allowed through approval code
         ElseIf Val(lsDetail) = 2 Then
            MsgBox "Selling of MC(s) without CSR needs approval code!", vbCritical, "Warning"
            
            'kalyptus - 2014.04.12 03:51pm
            'if below last price then asked for approval code
            If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
               MsgBox "User is not authorized to Approve MC Sales w/o CSR/PNP Clearane!" & vbCrLf & _
                     "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
               GoTo endProc
            End If
            
            If isValidApproveCode( _
               p_oSCA_Const.pxeMCSalesWoPNP, _
               p_oAppDrivr.BranchCode, _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               Master("sDRNoxxxx"), _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveName
               
               p_sSQLCSRx = "xTransNox"
               p_sSQLCSRx = "INSERT INTO xxxSCA_Usage" & _
                           " SET sTransNox = " & strParm(p_sSQLCSRx) & _
                              ", sApprCode = " & strParm(lsApprovedCD) & _
                              ", sApproved = " & strParm(lsApproveID) & _
                              ", sSystemCD = " & strParm(p_oSCA_Const.pxeMCSalesWoPNP) & _
                              ", sSourceNo = " & strParm(Master("sTransNox")) & _
                              ", sSourceCD = " & strParm(pxeMCSales) & _
                              ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If


      'Kalyptus - 2015.10.30
      'Integrate G_Card_Application Entry here
      Set p_oGCardAppl = Nothing
      'kalyptus - 2017.05.09 03:57pm
      'Allow condition that will not allow an entry of G_Card for MC Sales Replacement...
      If Trim(.Detail(0, "sSerialID")) <> "" And .Detail(0, "cMotorNew") = xeYes And Trim(.Detail(0, "sReplMCID")) = "" Then
         If MsgBox("Do you want to enter the GCard application for this MC Sales?", vbYesNo, "Confirmation") = vbYes Then
            If Not getCardApp() Then GoTo endProc
         End If
      ElseIf Trim(.Detail(0, "sSerialID")) <> "" And .Detail(0, "cMotorNew") = xeNo And InStr(1, p_oAppDrivr.getConfiguration("GCard2H"), p_oAppDrivr.BranchCode) > 0 Then
         'allow entry of gcard app for 2h branches only...
         If MsgBox("Do you want to enter the GCard application for this MC Sales?", vbYesNo, "Confirmation") = vbYes Then
            If Not getCardApp() Then GoTo endProc
         End If
      End If

      'get Assured Information for the Mafre Insurance
'      If IFNull(p_oAppDrivr.getConfiguration("cMAPFREIn", p_oAppDrivr.BranchCode), "0") = xeYes Then
'         If Trim(.Detail(0, "sSerialID")) <> "" And .Detail(0, "cMotorNew") = xeYes Then
'            If getAssuredInfo() = False Then GoTo endProc
'         End If
'      End If

      'Set client type here
      p_oClient.Master("cMCClient") = 1
      If .Master("cPaymForm") = "1" Or _
         .Master("cPaymForm") = "2" Then p_oClient.Master("cLRClient") = 1
      
      If p_oClient.SaveClient() = False Then
          MsgBox "Client Info not save!"
          GoTo endProc
      End If

      .Master("sClientID") = p_oClient.Master("sClientID")

      .Master("sCoCltID1") = ""
      p_sCoCltNm1 = ""

      If p_oCoClt1.Master("sLastName") <> "" Then
         p_oCoClt1.Master("cMCClient") = 1
         If .Master("cPaymForm") = "1" Or _
               .Master("cPaymForm") = "2" Then p_oCoClt1.Master("cLRClient") = 1
         
         If p_oCoClt1.SaveClient = False Then
            MsgBox "Co-client #1 info not save and will be remove!" & vbCrLf & _
                         "Please make sure to re-enter co-client info.", vbCritical, "Warning"
            GoTo endProc
         End If

         p_sCoCltNm1 = p_oCoClt1.Master("sLastName") + ", " + p_oCoClt1.Master("sFrstName") + " " + Trim(p_oCoClt1.Master("sSuffixNm")) + IIf(Trim(p_oCoClt1.Master("sSuffixNm")) = "", "", " ") + p_oCoClt1.Master("sMiddName")
         .Master("sCoCltID1") = p_oCoClt1.Master("sClientID")
      End If
      
      RaiseEvent MasterRetrieved(26)
      
      .Master("sCoCltID2") = ""
      p_sCoCltNm2 = ""
      
      If p_oCoClt2.Master("sLastName") <> "" Then
         p_oCoClt2.Master("cMCClient") = 1
         If .Master("cPaymForm") = "1" Or _
            .Master("cPaymForm") = "2" Then p_oCoClt2.Master("cLRClient") = 1
      
         If p_oCoClt2.SaveClient = False Then
            MsgBox "Co-client #2 info not save and will be remove!" & vbCrLf & _
                         "Please make sure to re-enter co-client info.", vbCritical, "Warning"
            GoTo endProc
         End If

         p_sCoCltNm2 = p_oCoClt2.Master("sLastName") + ", " + p_oCoClt2.Master("sFrstName") + " " + Trim(p_oCoClt2.Master("sSuffixNm")) + IIf(Trim(p_oCoClt2.Master("sSuffixNm")) = "", "", " ") + p_oCoClt2.Master("sMiddName")
         .Master("sCoCltID2") = p_oCoClt2.Master("sClientID")
      End If

      RaiseEvent MasterRetrieved(27)

      ' save credit application if one exists
      Select Case .Master("cPaymForm")
      Case "0"    ' Cash
         If .Master("nAmtPaidx") + p_nAdvPaymx < .Master("nTranTotl") Then
            If .Master("cTranType") = "1" Then
               MsgBox "Invalid amount paid!!!" & vbCrLf & _
                     "Paid amount must fit from payment type!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         
         'Application No should be empty
         .Master("sApplicNo") = ""
         If .Master("cTranType") = "1" Then
            If pbHotItemx Then
               If Not getHotItemApplication Then GoTo endProc
            End If
         End If
      
'         ' XerSys - 2013-05-08
'         '  Check given price if motorcycle sales
'         If Not (.Detail(0, "sSerialID") = "" Or .Detail(0, "cMotorNew") = xeNo) Then
'            If Not checkSRP(.Master("cPaymForm")) Then GoTo endProc
'         End If
      Case "1", "2"  ' Installment; Cash Balance
         ' XerSys - 2014-09-02
         '  Activate the following code and add a testing for a valid serial id
         ' XerSys - 2013-08-19
         '  Remove the following code because it was process in the getAccount procedure
         If .Master("sApplicNo") = Empty Or .Master("sApplicNo") = "N-O-N-E" Then
            If .Detail(0, "sSerialID") = Empty Then
               .Master("sApplicNo") = "N-O-N-E"
            ElseIf IsPromoRoute() Then
               .Master("sApplicNo") = "N-O-N-E"
            Else
               Call getApplication
            End If
         End If

         'Application No should NOT be empty
         If .Master("sApplicNo") = Empty Then
            MsgBox "Customer Has no Credit Application!"
            GoTo endProc
         End If
         
         If .Master("nAmtPaidx") + p_nAdvPaymx >= .Master("nTranTotl") Then
            MsgBox "Invalid amount paid!!!" & vbCrLf & _
               "Paid amount must fit from payment type!!!", vbCritical, "Warning"
            GoTo endProc
         End If

         ' @@@ save application
         If .Master("sApplicNo") <> "N-O-N-E" Then
            If p_oLRApplic.EditMode = xeModeUnknown Then
               If p_oLRApplic.OpenTransaction(.Master("sApplicNo")) = False Then GoTo endProc
               If p_oLRApplic.UpdateTransaction = False Then GoTo endProc
            ElseIf p_oLRApplic.EditMode = xeModeReady Then
               If p_oLRApplic.UpdateTransaction = False Then GoTo endProc
            Else
               p_oLRApplic.Master("sRemarksx") = "«Precreated : " & .Master("sTransNox") & "»"
            End If

            'We don't have update this info if just opened
            If p_oLRApplic.EditMode = xeModeAddNew Then
               p_oLRApplic.Master("sClientID") = .Master("sClientID")
               p_oLRApplic.Master("dAppliedx") = .Master("dTransact")
            End If

            If p_oLRApplic.SaveTransaction = False Then
               MsgBox "Unable to save Credit Application"
               GoTo endProc
            End If

            .Master("sApplicNo") = p_oLRApplic.Master("sTransNox")
         End If
         
'         ' XerSys - 2013-05-08
'         '  Check given price if motorcycle sales
'         If Not (.Detail(0, "sSerialID") = "" Or .Detail(0, "cMotorNew") = "") Then
'            If Not checkSRP(.Master("cPaymForm")) Then GoTo endProc
'         End If
      Case "3"    'Term
         'Application No should be empty
         .Master("sApplicNo") = ""
      Case Else   'Term With Registration/Financing
         If .Master("nAmtPaidx") + p_nAdvPaymx >= .Master("nTranTotl") Then
            MsgBox "Invalid amount paid!!!" & vbCrLf & _
               "Paid amount must fit from payment type!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         'Application No should be empty
         .Master("sApplicNo") = ""
      End Select
'900         If .Master("nAmtPaidx") > 0 Then
'910            .Master("cTranStat") = xeStateClosed
'920         Else
'930            .Master("cTranStat") = xeStateOpen
'940         End If

      'iMac [2015-12-28 11:45:00]
      If Not saveInquiry Then GoTo endProc

      'MS-2007.12.18
      If .EditMode = xeModeAddNew Then
         .Master("sTransNox") = GetNextCode(.MasterTable, "sTransNox", True, _
                                    p_oAppDrivr.Connection, True, p_sBranchCd)
      End If
      
      For pnCtr = 0 To .ItemCount - 1
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "nEntryNox") = pnCtr + 1
      Next
   End With

   Cancel = False

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
   GoTo endProc
End Sub

Private Function IsAutoQM() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "IsAutoQM"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   IsAutoQM = True
   
   lsSQL = "SELECT sTransNox" & _
          " FROM MC_LR_QuickMatch" & _
          " WHERE sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
          " LIMIT 1"
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If loRS.EOF Then IsAutoQM = False
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function

Private Function IsAssured() As Boolean
   Dim lsProcName As String
   
   lsProcName = "IsAssured"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function

Private Function IsPromoRoute() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "IsPromoRoute"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   IsPromoRoute = False
   
   lsSQL = "SELECT *" & _
          " FROM TownCity" & _
          " WHERE (sTownName like '%GGC%' OR sTownName like '%Guanzon%')" & _
          "   AND sTownIDxx = " & strParm(p_oClient.Master("sTownIDxx"))
   Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   If Not loRS.EOF Then IsPromoRoute = True
   
   'kalyptus - 2011.09.03
   'Include testing if employee
   If Not IsPromoRoute Then
      lsSQL = "SELECT sEmployID" & _
                   " FROM Employee_Master001" & _
                   " WHERE sEmployID = " & strParm(p_oClient.Master("sClientID"))
      Set loRS = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
            
            If Not loRS.EOF Then IsPromoRoute = True
         End If
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function

Private Function IsFirstPayModify(ByVal fdPurchase As Date, _
                                    ByVal fdFrstPaym As Date) As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   
   lsProcName = "IsFirstPayModify"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   IsFirstPayModify = False
   Select Case Day(fdPurchase)
   Case 1 To 3
      If Day(fdFrstPaym) <> 5 Then IsFirstPayModify = True
   Case 4 To 6
      If Day(fdFrstPaym) <> 8 Then IsFirstPayModify = True
   Case 7 To 9
      If Day(fdFrstPaym) <> 11 Then IsFirstPayModify = True
   Case 10 To 12
      If Day(fdFrstPaym) <> 14 Then IsFirstPayModify = True
   Case 13 To 15
      If Day(fdFrstPaym) <> 17 Then IsFirstPayModify = True
   Case 16 To 18
      If Day(fdFrstPaym) <> 20 Then IsFirstPayModify = True
   Case 19 To 21
      If Day(fdFrstPaym) <> 23 Then IsFirstPayModify = True
   Case 22 To 24
      If Day(fdFrstPaym) <> 26 Then IsFirstPayModify = True
   Case 25 To 28
      If Day(fdFrstPaym) <> 28 Then IsFirstPayModify = True
   Case 29 To 31
      If Day(fdFrstPaym) <> 3 Then IsFirstPayModify = True
   End Select
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "()"
   GoTo endProc
End Function


'kalyptus - 08/23/2012 03:02 pm
'Used in resetting the status of provisionary receipt
' XerSys - 2014-03-25
'  Reseting of advance payment must be through mc order object
Private Function resetAdvance() As Boolean
   Dim lsProcName As String
'   Dim lsSQL As String
'   Dim loRS As Recordset
   Dim lsTransNox As String
   Dim lnCashAmtx As Double
   Dim lnCheckAmt As Double
   Dim loMCOrder As clsMCCustOrder
   
   lsProcName = "resetAdvance"
   'On Error GoTo errProc
   
   With p_oAdvPaym
      .MoveFirst
      
      lsTransNox = .Fields("sTransNox")
      lnCashAmtx = 0
      lnCheckAmt = 0
      
      Set loMCOrder = New clsMCCustOrder
      With loMCOrder
         Set .AppDriver = p_oAppDrivr
         If Not .InitTransaction Then GoTo endProc
      End With
      
      Do Until .EOF
         If lsTransNox <> .Fields("sTransNox") Then
            If Not loMCOrder.DebitAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt) Then
               GoTo endProc
            End If
            
            lsTransNox = .Fields("sTransNox")
            lnCashAmtx = 0
            lnCheckAmt = 0
         End If
         
         If .Fields("xSrcTrans") = xeSrcOReceipt Then
            lnCashAmtx = .Fields("nCredtAmt")
         Else
            lnCheckAmt = .Fields("nCredtAmt")
         End If
         
         .MoveNext
      Loop
      If Not loMCOrder.DebitAdvPayment(lsTransNox, lnCashAmtx, lnCheckAmt) Then
         GoTo endProc
      End If
   End With
   
   resetAdvance = True
   
   ' XerSys 2014-03-26
   '  Disable the following codes
'   With p_oBaseClas
'      If .Master("sClientID") = Empty Then GoTo endWithWarning
'
'      lsSQL = "SELECT a.sTransNox" & _
'                  ", a.sPRNoxxxx" & _
'                  ", a.dTransact" & _
'                  ", a.nTranAmtx" & _
'                  ", a.cTranStat" & _
'                  ", b.sTransNox xChkRcvNo" & _
'               " FROM Provisionary_Receipt_Master a" & _
'                     " LEFT JOIN Checks_Received b" & _
'                        " ON a.sTransNox = b.sReferNox" & _
'                           " AND b.sSourceCd = " & strParm("PRec") & _
'               " WHERE a.sClientID = " & strParm(.Master("sClientID")) & _
'                  " AND a.sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
'                  " AND a.cTranType = " & strParm(0) & _
'                  " AND (a.cTranStat IN (" & strParm(xeStateClosed) & ", " & strParm(xeStatePosted) & ")" & _
'                   " AND a.sReferNox = " & strParm(.Master("sTransNox")) & _
'                   " AND a.sSourceCd = " & strParm("MCSl") & ")" & _
'               " ORDER BY sPRNoxxxx"
'      Set loRS = New Recordset
'      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
'      Set loRS.ActiveConnection = Nothing
'
'      Do Until loRS.EOF
'         lsSQL = "UPDATE Provisionary_Receipt_Master" & _
'                " SET cTranStat = " & strParm(xeStateOpen) & _
'                   ", sReferNox = ''" & _
'                   ", sSourceCD = ''" & _
'                " WHERE sTransNox = " & strParm(loRS("sTransNox"))
'         p_oAppDrivr.Execute lsSQL, "Provisionary_Receipt_Master"
'
'         ' test if advance payment is check
'         If Not IsNull(loRS("xChkRcvNo")) Then
'            lsSQL = "UPDATE Checks_Received" & _
'                     " SET cChckStat = " & strParm(xeStateOpen) & _
'                        ", cTranStat = " & strParm(xeStateOpen) & _
'                        ", cDepositd = " & strParm(xeNo) & _
'                        ", dDepositd = NULL" & _
'                        ", dStatChng = NULL" & _
'                     " WHERE sTransNox = " & strParm(loRS("xChkRcvNo"))
'
'            p_oAppDrivr.Execute lsSQL, "Checks_Received"
'         End If
'         loRS.MoveNext
'      Loop
'
'      lsSQL = "SELECT sTransNox" & _
'                  ", sORNoxxxx" & _
'                  ", dTransact" & _
'                  ", nTranAmtx" & _
'                  ", cTranStat" & _
'               " FROM Receipt_Master" & _
'               " WHERE sClientID = " & strParm(.Master("sClientID")) & _
'                  " AND sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
'                  " AND cTranType = " & strParm(0) & _
'                  " AND (cTranStat = " & strParm(xeStatePosted) & _
'                  " AND sReferNox = " & strParm(.Master("sTransNox")) & _
'                  " AND sSourceCd = " & strParm("MCSl") & ")" & _
'               " ORDER BY sORNoxxxx"
'
'      Set loRS = New Recordset
'      loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
'      Set loRS.ActiveConnection = Nothing
'
'      Do Until loRS.EOF
'         lsSQL = "UPDATE Receipt_Master" & _
'                " SET cTranStat = " & strParm(xeStateClosed) & _
'                   ", sReferNox = ''" & _
'                   ", sSourceCD = ''" & _
'                " WHERE sTransNox = " & strParm(loRS("sTransNox"))
'         p_oAppDrivr.Execute lsSQL, "Receipt_Master"
'
'         loRS.MoveNext
'      Loop
'   End With
   
endProc:
   Exit Function
endWithWarning:
   MsgBox "No Record Found for this Customer!" & _
               vbCrLf & " Can Not Process Advance Payment", vbCritical, "Warning"
   GoTo endProc
errProc:
   ShowError lsProcName
End Function

Private Function checkSRP(ByVal lcPaymForm As String) As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
   Dim lnRights As Integer
   Dim lsCodeType As String
   
   Dim lnDiscount As Double
   Dim lnMaxDiscx As Double

   lsProcName = "checkSRP"
   'On Error GoTo errProc
   
   p_sSQLSRPx = ""

   Select Case lcPaymForm
   Case "0", "1" ' Cash / Cash Balance
      If p_oBaseClas.Detail(0, "nUnitPrce") < p_oMCPrice.LastPrice Then
         lnDiscount = p_oMCPrice.LastPrice - p_oBaseClas.Detail(0, "nUnitPrce")
         
         MsgBox "Unit price is below the Suggested Retail Price for this Model!" & vbCrLf & _
                "This requires an approval from an authorized personnel..." & vbCrLf & vbCrLf & _
                "For additional detail, you can contact Guanzon Group's MIS/SSG.", vbOKOnly, "Validation"
         
         'kalyptus - 2014.04.12 03:51pm
         'if below last price then asked for approval code
         If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
            MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
                  "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
            GoTo endProc
         End If
         
         'iMac 2016.09.19
         'validate promo source
         Select Case p_oBaseClas.Master("cPromoTyp")
         Case "0"
            lsCodeType = p_oSCA_Const.pxeMCDiscount
         Case "2"
            lsCodeType = p_oSCA_Const.pxeAdditional
         Case "1"
            lsCodeType = p_oSCA_Const.pxeTeleMktg
         Case "3"
            lsCodeType = p_oSCA_Const.pxeBiyahingFiesta
         End Select
         
         Select Case p_oBaseClas.Master("cPromoTyp")
         Case "1" 'telemarketing
            lnMaxDiscx = CDbl(IFNull(p_oAppDrivr.getConfiguration("TLMSPAVl"), 0))
            
            If lnDiscount > lnMaxDiscx Then
               MsgBox "Discount given is greater than the maximum amount." & vbCrLf & vbCrLf & _
                        "Please ask assistance to MIS Department.", vbCritical, "Warning"
               GoTo endProc
            End If
         
            If isValidApproveCode( _
               lsCodeType, _
               "", _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               p_sFullName, _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveName
               
               lsSQL = "xTransNox"
               p_sSQLGvwy = "INSERT INTO xxxSCA_Usage" & _
                           " SET sTransNox = " & strParm(lsSQL) & _
                              ", sApprCode = " & strParm(lsApprovedCD) & _
                              ", sApproved = " & strParm(lsApproveID) & _
                              ", sSystemCD = " & strParm(lsCodeType) & _
                              ", sSourceNo = " & strParm(Master("sTransNox")) & _
                              ", sSourceCD = " & strParm(pxeMCSales) & _
                              ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                              ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
         Case 0 'NONE or with SPA
            'mac 2023.10.03
            '  if general manager code/grace padlan, remove limit on discount
            If Mid(lsApprovedCD, 4, 1) <> "9" Then
               lnMaxDiscx = CDbl(IFNull(p_oAppDrivr.getConfiguration("AHSPAVal"), 0))
            
               If lnDiscount > lnMaxDiscx Then
                  MsgBox "Discount given is greater than the maximum amount." & vbCrLf & vbCrLf & _
                           "Please ask assistance to MIS Department.", vbCritical, "Warning"
                  GoTo endProc
               End If
            End If
         
            If isValidApproveCode( _
               p_oSCA_Const.pxeMCDiscount, _
               p_oAppDrivr.BranchCode, _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               Master("sDRNoxxxx"), _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveName

               lsSQL = "xTransNox"
               p_sSQLSRPx = "INSERT INTO xxxSCA_Usage" & _
                      " SET sTransNox = " & strParm(lsSQL) & _
                         ", sApprCode = " & strParm(lsApprovedCD) & _
                         ", sApproved = " & strParm(lsApproveID) & _
                         ", sSystemCD = " & strParm(p_oSCA_Const.pxeMCDiscount) & _
                         ", sSourceNo = " & strParm(Master("sTransNox")) & _
                         ", sSourceCD = " & strParm(pxeMCSales) & _
                         ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                         ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
         End Select
      End If
   Case "2" ' Installment
      If p_oBaseClas.Detail(0, "nUnitPrce") < p_oMCPrice.SelPrice Then
         MsgBox "Unit price is below the Suggested Retail Price for this Model!" & vbCrLf & _
               "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
         GoTo endProc
      End If
      
      'mac 2020.09.05
      '  validate down payment regardless of sold status
      
      'Check Down Payment vs Minimum Down if New
      'If p_oBaseClas.Detail(0, "cMotorNew") = "1" Then
         If (p_oLRMaster.Master("nDownPaym")) < p_oMCPrice.MinimumDown Then
            MsgBox "Down Payment is below the Minimum Down for this Model!" & vbCrLf & _
                  "Approval Code is Required!!!", vbInformation, "Notice"
                  
            If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
               MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
                     "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
               GoTo endProc
            End If
            
            If isValidApproveCode( _
               p_oSCA_Const.pxeMCDownpayment, _
               p_oAppDrivr.BranchCode, _
               Mid(lsApprovedCD, 4, 1), _
               Master("dTransact"), _
               Master("sDRNoxxxx"), _
               lsApprovedCD) Then
               
               p_oBaseClas.Master("sApproved") = lsApproveName

               lsSQL = "xTransNox"
               p_sSQLSRPx = "INSERT INTO xxxSCA_Usage" & _
                      " SET sTransNox = " & strParm(lsSQL) & _
                         ", sApprCode = " & strParm(lsApprovedCD) & _
                         ", sApproved = " & strParm(lsApproveID) & _
                         ", sSystemCD = " & strParm(p_oSCA_Const.pxeMCDownpayment) & _
                         ", sSourceNo = " & strParm(Master("sTransNox")) & _
                         ", sSourceCD = " & strParm(pxeMCSales) & _
                         ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                         ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
                  
            Else
               MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
                  "Verify entry then try again!", vbCritical, "Warning"
               GoTo endProc
            End If
            'GoTo endProc
         End If
      'End If
   End Select
   
   checkSRP = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

' XerSys - 2016-09-23
'  Validate SRP of 2H motorcycle
Private Function valid2HPrice() As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
   Dim lnRights As Integer
   Dim lsCodeType As String
   Dim lasClass() As String
   Dim lanQty() As Integer
   Dim lnCtr As Integer

   lsProcName = "valid2HPrice"
   'On Error GoTo errProc
   
   p_sSQLSRPx = ""
   
   With p_oBaseClas
      valid2HPrice = True
      For lnCtr = 0 To .ItemCount - 1
         If Not isDiscOK(.Detail(lnCtr, "sSerialID"), .Detail(lnCtr, "nUnitPrce")) Then
            valid2HPrice = False
            Exit For
         End If
      Next
   End With
   
   If Not valid2HPrice Then
      MsgBox "Unit price is below the Suggested Retail Price for this Model!" & vbCrLf & _
             "This requires an approval from an authorized personnel..." & vbCrLf & vbCrLf & _
             "For additional detail, you can contact Guanzon Group's MIS/SSG.", vbOKOnly, "Validation"
      
      'kalyptus - 2014.04.12 03:51pm
      'if below last price then asked for approval code
      If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
         MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
               "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
         GoTo endProc
      End If
      
      If isValidApproveCode( _
         p_oSCA_Const.pxeMCDiscount, _
         p_oAppDrivr.BranchCode, _
         Mid(lsApprovedCD, 4, 1), _
         Master("dTransact"), _
         Master("sDRNoxxxx"), _
         lsApprovedCD) Then
         
         p_oBaseClas.Master("sApproved") = lsApproveName

         lsSQL = "xTransNox"
         p_sSQLSRPx = "INSERT INTO xxxSCA_Usage" & _
                " SET sTransNox = " & strParm(lsSQL) & _
                   ", sApprCode = " & strParm(lsApprovedCD) & _
                   ", sApproved = " & strParm(lsApproveID) & _
                   ", sSystemCD = " & strParm(p_oSCA_Const.pxeMCDiscount) & _
                   ", sSourceNo = " & strParm(Master("sTransNox")) & _
                   ", sSourceCD = " & strParm(pxeMCSales) & _
                   ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                   ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            
      Else
         MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
            "Verify entry then try again!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   valid2HPrice = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

'  mac 2020.08.24
'     validate commision
Private Function validSalesCommision() As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsApprovedCD As String, lsApproveID As String, lsApproveName As String
   Dim lnRights As Integer
   Dim lsCodeType As String
   Dim lasClass() As String
   Dim lanQty() As Integer
   Dim lnCtr As Integer
   Dim lnMaxCommision As Integer

   lsProcName = "validSalesCommision"
   'On Error GoTo errProc
   
   'mac 2020-08-24
   '  do not allow a HIM with commision
   If pbHotItemx = True And p_nCommAmtx > 0# Then
      MsgBox "Hot Item Models with commsion is not allowed.", vbCritical, "Warning"
      GoTo endProc
   End If
   '  end - do not allow a HIM with commision
   
   'mac 2020-08-24
   '  commision must be only less than or equal to amount set on config
   lnMaxCommision = IFNull(p_oAppDrivr.getConfiguration("MaxComm"), 0)
   
   If lnMaxCommision <= 0 And p_nCommAmtx > 0# Then
      MsgBox "Commision is UNSET or NOT ALLOWED.", vbInformation, "Notice"
      GoTo endProc
   End If
   
   If p_nCommAmtx > lnMaxCommision Then
      MsgBox "Commission exceeds the allowable amount.", vbCritical, "Warning"
      GoTo endProc
   End If
   '  end - commision must be only less than or equal to amount set on config
   
   p_sSQLComm = ""
   
   If p_nCommAmtx > 0# Then
      MsgBox "Sales commission has been detected!" & vbCrLf & _
             "This requires an approval from an authorized personnel..." & vbCrLf & vbCrLf & _
             "For additional detail, you can contact Guanzon Group's MIS/SSG.", vbOKOnly, "Validation"
      
      If Not GetCodeApproval(p_oAppDrivr, lsApprovedCD, lsApproveID, lsApproveName) Then
         MsgBox "User is not authorized to Approve such Discount!" & vbCrLf & _
               "Verify your entry then try again!", vbOKOnly + vbCritical, "Warning"
         GoTo endProc
      End If
      
      If isValidApproveCode( _
         p_oSCA_Const.pxeSalesCommision, _
         p_oAppDrivr.BranchCode, _
         Mid(lsApprovedCD, 4, 1), _
         Master("dTransact"), _
         Master("sDRNoxxxx"), _
         lsApprovedCD) Then
         
         p_oBaseClas.Master("sApproved") = lsApproveName

         lsSQL = "xTransNox"
         p_sSQLComm = "INSERT INTO xxxSCA_Usage" & _
                " SET sTransNox = " & strParm(lsSQL) & _
                   ", sApprCode = " & strParm(lsApprovedCD) & _
                   ", sApproved = " & strParm(lsApproveID) & _
                   ", sSystemCD = " & strParm(p_oSCA_Const.pxeSalesCommision) & _
                   ", sSourceNo = " & strParm(Master("sTransNox")) & _
                   ", sSourceCD = " & strParm(pxeMCSales) & _
                   ", sModified = " & strParm(p_oAppDrivr.UserID) & _
                   ", dModified = " & dateParm(p_oAppDrivr.ServerDate)
            
      Else
         MsgBox "Invalid APPROVAL CODE detected." & vbCrLf & _
            "Verify entry then try again!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   validSalesCommision = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

'mac 2022.09.21
Private Function isDiscOK(ByVal lsSerialID As String, ByVal lnUnitPrce As Double) As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lnDiscountd As Double
   
   lsProcName = "isDiscOK"
   'On Error GoTo errProc
   
   lsSQL = "SELECT" & _
               "  a.sSerialID" & _
               ", a.sModelIDx" & _
               ", b.nPOwnSRPx" & _
            " FROM" & _
              " MC_Serial a" & _
              " LEFT JOIN MC_Model d" & _
                  " ON a.sModelIDx = d.sModelIDx" & _
              " LEFT JOIN MC_Serial_Classification_New b" & _
                  " ON a.sSerialID = b.sSerialID" & _
            " WHERE a.sSerialID = " & strParm(lsSerialID) & _
              " AND a.cSoldStat = '1'"
            
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then GoTo endProc
   
   If IsNull(loRS("nPOwnSRPx")) Then
      isDiscOK = True
      GoTo endProc
   End If
   
   If lnUnitPrce < loRS("nPOwnSRPx") Then
      GoTo endProc
   End If
   
   isDiscOK = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

' XerSys - 2015-05-11
'  The following codes below were added for the agent field
Private Function initAgent() As Boolean
   Dim lsProcName As String
   
   lsProcName = "initAgent"
   'On Error GoTo errProc
   
   Set p_oAgent = New clsNeoClient
   With p_oAgent
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oAgentInfo = New Recordset
   p_oAgentInfo.Open AddCondition(p_sAgentSQL, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   p_oAgentInfo.AddNew
   Call initAgentInfo
   Call initAgentTrans
   
   p_bNewAgent = True
   
   initAgent = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

Private Function initAgentInfo() As Boolean
   With p_oAgentInfo
      ' assign default values
      .Fields("sClientID") = ""
      .Fields("sBranchCD") = ""
      .Fields("nNoofDepn") = 0
      .Fields("nNoofChld") = 0
      .Fields("sEmployID") = ""
      .Fields("sRelation") = ""
      .Fields("nNoOfCust") = 0
      .Fields("nCommissn") = 0
      .Fields("dCltSince") = Date
      .Fields("sOthersxx") = ""
      .Fields("cRecdStat") = xeRecStateActive
   End With
End Function

Private Sub initAgentTrans()
   p_sAgentNme = ""
   p_nCommAmtx = 0
   p_sReferNox = ""
   p_cReleased = xeNo
   p_sReleased = ""
   p_dReleased = Date
End Sub

Private Function getAgent(ByVal lsField As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient
   Dim loAgent As frmMC_Agent
   
   lsProcName = "getAgent"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
      'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   If lsField <> "" Then
      If Trim(LCase(lsField)) = Trim(LCase(p_sAgentNme)) Then GoTo endProc
      lbExist = loClient.SearchClient(lsField, False)
   Else
      GoTo endWithClear
   End If

   If Not lbExist Then
      lasName = GetSplitedName(lsField)
      loClient.Master("sLastName") = lasName(0)
      loClient.Master("sFrstName") = lasName(1)
   End If

   If loClient.getClient Then
      Set p_oAgent = loClient
   End If

   p_sAgentNme = p_oAgent.Master("sLastName") + ", " + p_oAgent.Master("sFrstName") + _
                  " " + Trim(p_oAgent.Master("sSuffixNm")) + _
                  IIf(Trim(p_oAgent.Master("sSuffixNm")) = "", "", " ") + _
                  p_oAgent.Master("sMiddName")
   
   
   Call loadAgent(p_oAgent.Master("sClientID"))
   
   Set loAgent = frmMC_Agent
   With loAgent
      Set .AppDriver = p_oAppDrivr
      .AgentID = p_oAgent.Master("sClientID")
      .AgentName = p_sAgentNme
      .AgentAddress = p_oAgent.Master("sAddressx")
      
      .NoOfDependent = p_oAgentInfo("nNoofDepn")
      .NoOfChild = p_oAgentInfo("nNoofChld")
      .EmployeeCd = p_oAgentInfo("sEmployID")
      .RelationCd = p_oAgentInfo("sRelation")
      .Commission = p_nCommAmtx
      .ClientSince = IFNull(p_oAgentInfo("dCltSince"), Date)
      .OtherInfo = p_oAgentInfo("sOthersxx")
      .GCardHolder = p_bGCardAgent
      .Show 1
      
      If .Cancelled Then GoTo endProc
      
      p_oAgentInfo("sClientID") = p_oAgent.Master("sClientID")
      p_oAgentInfo("nNoofDepn") = .NoOfDependent
      p_oAgentInfo("nNoofChld") = .NoOfChild
      p_oAgentInfo("sEmployID") = .EmployeeCd
      p_oAgentInfo("sRelation") = .RelationCd
      p_nCommAmtx = .Commission
      p_oAgentInfo("dCltSince") = .ClientSince
      p_oAgentInfo("sOthersxx") = .OtherInfo
   End With
   getAgent = True
   
endProc:
   RaiseEvent MasterRetrieved(31)
   Exit Function
endWithClear:
   Set p_oAgent = New clsNeoClient
   p_sAgentNme = ""
   GoTo endProc
errProc:
    ShowError lsProcName & "( " & lsField _
                        & ", " & lbSearch & " )"
End Function

Private Function loadAgentTrans() As Boolean
   Dim lsProcName As String
   Dim loRS As Recordset
   Dim lsSQL As String
   
   lsProcName = "loadAgentTrans"
   'On Error GoTo errProc
   
   Set loRS = New Recordset
   loRS.Open "SELECT * FROM MC_SO_Agent" & _
               " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")), _
                     p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then
      ' No agent record
      Call initAgentTrans
   Else
      p_nCommAmtx = loRS("nCommAmtx")
      p_sReferNox = loRS("sReferNox")
      p_cReleased = loRS("cReleased")
      p_sReleased = loRS("sReleased")
      p_dReleased = IFNull(loRS("dReleased"), Date)
      
      Call loadAgent(loRS("sAgentIDx"))
      Call p_oAgent.OpenClient(loRS("sAgentIDx"))
   End If
   
   loadAgentTrans = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function loadAgent(ByVal lsAgentID As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsProcName = "loadAgent"
   'On Error GoTo errProc
   
   Set p_oAgentInfo = New Recordset
   lsSQL = AddCondition(p_sAgentSQL, "sClientID = " & strParm(lsAgentID))
   Debug.Print lsSQL
   p_oAgentInfo.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   'Mac 2018-06-01
   '  check first if the agent was a GCard Holder.
   lsSQL = "SELECT sClientID" & _
            " FROM G_Card_Master" & _
            " WHERE cCardStat = '4'" & _
               " AND sClientID = " & strParm(lsAgentID)
            
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   p_bGCardAgent = loRS.RecordCount <> 0
   Set loRS = Nothing
   ''''''
                     
   If p_oAgentInfo.EOF Then
      p_oAgentInfo.AddNew
      Call initAgentInfo
      
      p_bNewAgent = True
      GoTo endProc
   Else
      p_bNewAgent = False
   End If
   
   
   
   Call p_oAgent.OpenClient(p_oAgentInfo("sClientID"))
   p_sAgentNme = p_oAgent.Master("sLastName") + ", " + p_oAgent.Master("sFrstName") + _
                  " " + Trim(p_oAgent.Master("sSuffixNm")) + _
                  IIf(Trim(p_oAgent.Master("sSuffixNm")) = "", "", " ") + _
                  p_oAgent.Master("sMiddName")
                  
   loadAgent = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & lsAgentID & " )"
End Function

Private Function updateAgent() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "updateAgent"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If p_sAgentNme = "" Then
      updateAgent = True
      GoTo endProc
   End If
   
   If Not p_oAgent.SaveClient Then GoTo endProc
   
   If p_bNewAgent Then
      p_oAgentInfo("sBranchCd") = p_oAppDrivr.BranchCode
      p_oAgentInfo("dCltSince") = Master("dTransact")
      p_oAgentInfo("nCommissn") = p_nCommAmtx
      p_oAgentInfo("sClientID") = p_oAgent.Master("sClientID")
      lsSQL = ADO2SQL(p_oAgentInfo, "MC_Agent", , , p_oAppDrivr.ServerDate)
   Else
      lsSQL = ADO2SQL(p_oAgentInfo, "MC_Agent", "sClientID = " & strParm(p_oAgentInfo("sClientID")), , p_oAppDrivr.ServerDate)
   End If
   If lsSQL <> "" Then
      If p_oAppDrivr.Execute(lsSQL, "MC_Agent") <= 0 Then
         MsgBox "Unable to Create MC Agent Record!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If

   lsSQL = "INSERT INTO MC_SO_Agent" & _
            " SET sTransNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
               ", sAgentIDx = " & strParm(p_oAgentInfo("sClientID")) & _
               ", nCommAmtx = " & p_nCommAmtx & _
               ", sRefernox = " & strParm("") & _
               ", cGCRefrlx = " & strParm(IIf(p_bGCardAgent = True, "1", "0")) & _
               ", cReleased = " & strParm(xeNo) & _
               ", sReleased = " & strParm("") & _
               ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
            " ON DUPLICATE KEY" & _
            " UPDATE sAgentIDx = " & strParm(p_oAgentInfo("sClientID")) & _
               ", nCommAmtx = " & p_nCommAmtx & _
               ", sRefernox = " & strParm(p_sReferNox) & _
               ", cGCRefrlx = " & strParm(p_bGCardAgent) & _
               ", cReleased = " & strParm(p_cReleased) & _
               ", sReleased = " & strParm(p_sReleased)

   If p_oAppDrivr.Execute(lsSQL, "MC_SO_Agent") <= 0 Then
      MsgBox "Unable to Create MC Sales Agent Record!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   updateAgent = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function commisionReleased() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   
   commisionReleased = False
   
   lsSQL = "SELECT * FROM MC_SO_Agent" & _
            " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))
            
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If loRS.EOF Then
      GoTo endProc
   Else
      If loRS("cReleased") = xeNo Then
         GoTo endProc
      Else
         commisionReleased = True
      End If
   End If
   
endProc:
   Exit Function
End Function

Private Function saveInquiry() As Boolean
   Dim lsSQL As String
   Dim lsTransNox As String
   Dim lsProcName As String
   
   lsProcName = "saveInquiry"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'On Error GoTo errProc
   
   If p_oBaseClas.EditMode <> xeModeAddNew Then
      saveInquiry = True
      GoTo endProc
   End If
   
   lsTransNox = Inquiry
   
   If lsTransNox = "" Then
      saveInquiry = True
      GoTo endProc
   End If
   
   lsSQL = "UPDATE MC_Product_Inquiry" & _
            " SET dPurchase = " & dateParm(p_oBaseClas.Master("dTransact")) & _
               ", cTranStat = " & strParm(xeStatePosted) & _
            " WHERE sTransNox = " & strParm(lsTransNox)

   If p_oAppDrivr.Execute(lsSQL, "MC_Product_Inquiry") <= 0 Then
      MsgBox "Unable to Update Motorcycle Inquiry." & vbCrLf & _
               "Please Inform MIS Dept.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   saveInquiry = True
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function Inquiry() As String
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lsTableNme As String
   Dim lsCondition As String
   Dim loRS As Recordset
   
   lsProcName = "hasInquiry"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'On Error GoTo errProc
   lsTableNme = "MC_Product_Inquiry"
   
   lsSQL = "SELECT" & _
               "  sTransNox" & _
               ", dTransact" & _
               ", sClientID" & _
               ", sModelIDx" & _
               ", sColorIDx" & _
               ", dTargetxx" & _
               ", dPurchase" & _
               ", cPurcType" & _
               ", cTranStat" & _
            " FROM " & lsTableNme
   lsSQL = lsSQL & _
            " WHERE sClientID = " & strParm(p_oBaseClas.Master("sClientID")) & _
               " AND sModelIDx = " & strParm(p_oBaseClas.Detail(0, "sModelIDx")) & _
               " AND sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%") & _
               " AND cTranStat NOT IN ('2', '3')" & _
            " ORDER BY sTransNox"
   
   Set loRS = New Recordset
   
   lsCondition = AddCondition(lsSQL, "sColorIDx = " & strParm(IFNull(p_oBaseClas.Detail(0, "sColorIDx"), "")))
   loRS.Open lsCondition, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If loRS.EOF Then
      loRS.Close
      loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      Set loRS.ActiveConnection = Nothing
   
      If loRS.EOF Then
         Inquiry = ""
      Else
         Inquiry = loRS("sTransNox")
      End If
   Else
      Inquiry = loRS("sTransNox")
   End If
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function delAgent() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "delAgent"
   Debug.Print pxeMODULENAME & "." & lsProcName
   'On Error GoTo errProc
   
   If p_sAgentNme = "" Then
      delAgent = True
      GoTo endProc
   End If
   
   If p_oAgentInfo("nNoOfCust") = 0 Then
      lsSQL = "DELETE FROM MC_Agent" & _
               " WHERE sClientID = " & strParm(p_oAgentInfo("sClientID"))
      
      If p_oAppDrivr.Execute(lsSQL, "MC_Agent") <= 0 Then
         MsgBox "Unable to Delete MC Agent Record!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   lsSQL = "DELETE FROM MC_SO_Agent" & _
            " WHERE sTransNox = " & strParm(p_oBaseClas.Master("sTransNox"))

   If p_oAppDrivr.Execute(lsSQL, "MC_SO_Agent") <= 0 Then
      MsgBox "Unable to Delete MC Sales Agent Record!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   delAgent = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & " )"
End Function

Private Function delAssuredInfo() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim loOReceipt As clsORReceipt
   Dim lsClientID As String
   
   If IFNull(p_oBaseClas.Master("xInsurTrn"), "") = "" Then
      delAssuredInfo = True
      GoTo endProc
   End If
   
   Set loRS = New Recordset
   loRS.Open "SELECT sTransNox, cTranStat, sReferNox" & _
               " FROM Receipt_Master" & _
               " WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                  " AND sSourceCd = 'MCSl'" & _
                  " AND sSystemCd = 'MC'" & _
                  " AND cRegisFrm = '6' " & _
                  " AND sInsTypID = '0109'" & _
                  " AND cTranStat <> '3' " _
   , p_oAppDrivr.Connection, , , adCmdText
   
   If Not loRS.EOF Then
      Set loOReceipt = New clsORReceipt
      Set loOReceipt.AppDriver = p_oAppDrivr
      loOReceipt.Branch = p_sBranchCd
      loOReceipt.Parent = "clsMCSales"
               
      lsClientID = p_oBaseClas.Master("sClientID")
      loOReceipt.Filter = "a.sClientID = " & strParm(lsClientID)
         
      If loOReceipt.InitTransaction = False Then GoTo endProc
      If loOReceipt.SearchByRef(p_oBaseClas.Master("sTransNox"), "MCSl") = False Then
         MsgBox "Unable to Load Receipt Info!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If loOReceipt.DeleteTransaction = False Then
         MsgBox "Unable to Delete MAPFRE OR!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   lsSQL = "DELETE FROM MC_Accident_Insurance" & _
            " WHERE sSourceNo = " & strParm(p_oBaseClas.Master("sTransNox")) & _
               " AND sSourceCd = 'MCSl'"
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance") = 0 Then
      MsgBox "Unable to Delete Insurance Master.", vbCritical, "Warning"
      delAssuredInfo = False
      Exit Function
   End If
   
   lsSQL = "DELETE FROM MC_Accident_Insurance_Beneficiaries" & _
            " WHERE sTransNox = " & strParm(p_oBaseClas.Master("xInsurTrn"))
   Call p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance")
   
   delAssuredInfo = True
   
endProc:
   Exit Function
End Function

Private Function cancelAssuredInfo() As Boolean
   Dim loOReceipt As clsORReceipt
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsClientID As String
   Dim lsTransNox As String
   
   If IFNull(p_oBaseClas.Master("xInsurTrn"), "") = "" Then
      cancelAssuredInfo = True
      GoTo endProc
   End If
   
   Set loRS = New Recordset
   loRS.Open "SELECT sTransNox, cTranStat, sReferNox" & _
               " FROM Receipt_Master" & _
               " WHERE sReferNox = " & strParm(p_oBaseClas.Master("sTransNox")) & _
                  " AND sSourceCd = 'MCSl'" & _
                  " AND sSystemCd = 'MC'" & _
                  " AND cRegisFrm = '6' " & _
                  " AND sInsTypID = '0109'" & _
                  " AND cTranStat <> '3' " _
   , p_oAppDrivr.Connection, , , adCmdText
   
   If Not loRS.EOF Then
      Set loOReceipt = New clsORReceipt
      Set loOReceipt.AppDriver = p_oAppDrivr
      loOReceipt.Branch = p_sBranchCd
      loOReceipt.Parent = "clsMCSales"
               
      lsClientID = p_oBaseClas.Master("sClientID")
      loOReceipt.Filter = "a.sClientID = " & strParm(lsClientID)
         
      If loOReceipt.InitTransaction = False Then GoTo endProc
      If loOReceipt.SearchByRef(p_oBaseClas.Master("sTransNox"), "MCSl") = False Then
         MsgBox "Unable to Load Receipt Info!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If loOReceipt.CancelTransaction = False Then
         MsgBox "Unable to Cancel MAPFRE Receipt!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
         
   lsSQL = "UPDATE MC_Accident_Insurance" & _
            " SET cTranStat = " & xeStateCancelled & _
            " WHERE sSourceNo = " & strParm(p_oBaseClas.Master("sTransNox")) & _
               " AND sSourceCd = 'MCSl'"
            
   If p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance") = 0 Then
      MsgBox "Unable to Cancel Insurance Master.", vbCritical, "Warning"
      cancelAssuredInfo = False
      Exit Function
   End If
   
   cancelAssuredInfo = True
endProc:
   Exit Function
End Function


Private Function getAssuredInfo() As Boolean
   Dim loForm As frmAssuredInfo
   
   Set loForm = New frmAssuredInfo
   
   With loForm
      Set .AppDriver = p_oAppDrivr
      Set .Client = p_oClient
      .ClientID = p_oBaseClas.Master("sClientID")
      .TransNox = p_oBaseClas.Master("sTransNox")
      
      If p_oBaseClas.Master("cPromoTyp") = "1" Then
         'iMac 2017.04.10
         '.TransNox = p_oBaseClas.Master("sTransNox") & "TLM"
         
         'iMac 2018.12.11
         If p_oBaseClas.Master("cPaymForm") = "0" Then 'cash
            .TransNox = p_oBaseClas.Master("sTransNox") & "TLMc"
         ElseIf p_oBaseClas.Master("cPaymForm") = "2" Then 'installment
            .TransNox = p_oBaseClas.Master("sTransNox") & "TLMi"
         End If
      End If
      
      .SaleDate = p_oBaseClas.Master("dTransact")
      .MCSerial = p_oBaseClas.Detail(0, "sSerialID")
      .ORNo = p_oBaseClas.Master("sORNoxxxx")
      .Parent = pxeMODULENAME
      .ViewOnly = False
      .Show vbModal
         
      If .Cancelled Then
         MsgBox "MAPFRE Insurance entry was disregarded!!!", vbInformation, "Notice"
      
         Set p_oAccident = Nothing
      Else
         Set p_oAccident = .Insurance
      End If
      
      getAssuredInfo = True
   End With
endProc:
   Set loForm = Nothing
   Exit Function
End Function

Private Function saveAssuredInfo() As Boolean
   saveAssuredInfo = True
   
   If p_oAccident Is Nothing Then
      Exit Function
   End If
   
   p_oAccident.Master("sSourceNo") = p_oBaseClas.Master("sTransNox")
   
   'kalyptus - 2016.06.17 04:36
   're-assigned the client id before calling the save procedure of the insurance object
   p_oAccident.Master("sClientID") = p_oBaseClas.Master("sClientID")
   saveAssuredInfo = p_oAccident.SaveTransaction
End Function

Function LoadInsurance() As Boolean
   Dim loForm As frmAssuredInfo
   Dim loInsurance As clsAccidentInsurance
   
   Set loForm = New frmAssuredInfo
   
   If IFNull(p_oBaseClas.Master("sClientID")) = "" Then GoTo endProc
   If p_oBaseClas.Detail(0, "cMotorNew") = xeNo Then GoTo endProc
'   If IFNull(p_oBaseClas.Master("xInsurTrn"), "") = "" Then GoTo endProc
   
   With loForm
      Set .AppDriver = p_oAppDrivr
      .TransNox = p_oBaseClas.Master("sTransNox")
      
'      p_oAccident.Master("sSourceNo") = p_oBaseClas.Master("sTransNox") & "TLM"
      '.ClientID = p_oBaseClas.Master("sClientID")
      .SaleDate = p_oBaseClas.Master("dTransact")
      .MCSerial = p_oBaseClas.Detail(0, "sSerialID")
      
      Set .Client = p_oClient
      
      .Parent = pxeMODULENAME
      .showMessage = True
      .ViewOnly = True
      .Show vbModal
         
      If Not .Cancelled Then
         Set loInsurance = .Insurance
         LoadInsurance = loInsurance.SaveTransaction
         GoTo endProc
      End If
         
      LoadInsurance = True
   End With
   
endProc:
   Exit Function
End Function

'Integrate G_Card_Application entry here...
Private Function getCardApp() As Boolean
   Dim lsSQL As String
   
   Dim loFrm As frmGCardApplication
   
   If p_oGCardAppl Is Nothing Then
      lsSQL = "SELECT" & _
                     "  sTransNox" & _
                     ", dTransact" & _
                     ", sSourceCd" & _
                     ", sModelIDx" & _
                     ", cPurcMode" & _
                     ", nAcctTerm" & _
                     ", sClientID" & _
                     ", sSrceInfo" & _
                     ", sReasonsx" & _
                     ", cCardType" & _
                     ", cApplType" & _
                     ", sPrevGCrd" & _
                     ", sSourceNo" & _
                     ", nAmtPaidx" & _
                     ", sNmOnCard" & _
                     ", sCompnyID" & _
                     ", sSoldByxx" & _
                     ", sSerialID"
      lsSQL = lsSQL & _
                     ", nYellowxx" & _
                     ", nWhitexxx" & _
                     ", nPointsxx" & _
                     ", cDigitalx" & _
                     ", cTranStat" & _
                     ", sModified" & _
                     ", dModified" & _
               " FROM G_Card_Application" & _
               " WHERE 0=1"
      Set p_oGCardAppl = New Recordset
      With p_oGCardAppl
         .Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
         .ActiveConnection = Nothing
         
         .AddNew
         .Fields("sTransNox") = GetNextCode("G_Card_Application", "sTransNox", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)
         'insert a condition for new/repo
         .Fields("sSourceCD") = IIf(p_oBaseClas.Detail(0, "cMotorNew") = xeYes, "M02910000005", "M02910000012")   'Please check G_Card_Points_Basis
         .Fields("sModelIDx") = ""
         .Fields("dTransact") = p_oBaseClas.Master("dTransact")
         'Cash = 0; Installment = 1
         .Fields("cPurcMode") = IIf(p_oBaseClas.Master("cPaymForm") = "1" Or p_oBaseClas.Master("cPaymForm") = "2", "1", "0")
         
         .Fields("sSrceInfo") = "000000000000000"
         .Fields("sReasonsx") = "0000000000"
         
         .Fields("cCardType") = "0"
         .Fields("cApplType") = "1" '1 means new

         .Fields("sSourceNo") = p_oBaseClas.Master("sDRNoxxxx")
         .Fields("nAmtPaidx") = 0
         .Fields("sCompnyID") = p_oAppDrivr.BranchCode
         .Fields("sSerialID") = p_oBaseClas.Detail(0, "sSerialID")
         .Fields("nYellowxx") = 0
         .Fields("nWhitexxx") = 0
         .Fields("nPointsxx") = 0
         .Fields("cDigitalx") = 0
         .Fields("cTranStat") = "0"
      End With
   End If
   
   Set loFrm = New frmGCardApplication
   Set loFrm.Master = p_oGCardAppl
   Set loFrm.AppDriver = p_oAppDrivr
   Set loFrm.MCSales = Me
   
   'Client info is reset to customer id
   p_oGCardAppl.Fields("sClientID") = p_oBaseClas.Master("sClientID")
   Set loFrm.Client = p_oClient
      
   loFrm.Show 1
   
   If Not loFrm.IsOkey Then Exit Function
   
   'mac 2019-07-09
   'get application card type
   p_oGCardAppl("cDigitalx") = loFrm.IsDigital
            
   'kalyptus - 2016.06.17 04:59pm
   'Set the information of the GCard Client to the GCard Client Variable
   'Note: the value of the variable is Nothing if gcard client and mc client is the same
   Set p_oGCCltx = loFrm.GCard_Client
            
   getCardApp = True
End Function

Public Function StringToHex(ByVal StrToHex As String) As String
   Dim strTemp   As String
   Dim strReturn As String
   Dim I As Long
   
   For I = 1 To Len(StrToHex)
       strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
       If Len(strTemp) = 1 Then strTemp = "0" & strTemp
       strReturn = strReturn & strTemp
   Next I
   StringToHex = strReturn
End Function

Private Function newcard(ByVal fsValue) As Integer
   newcard = RMJExecute(pxeJavaPath & "newcard.bat " & fsValue)
End Function

'mac 2020.09.02
'  RATEx = -2
'  Price = BN SRP - (BN SRP * (CLASS MIN RATE + RATEx) / 100)
'
'  where RATEx is the specified rate by the commitee to deduct on the mimimum rate for 2H classification
Private Function get2HPrice(ByVal lsSerialID As String) As Double
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lsProcName As String
   Dim lnMinRatex As Double
   
   lsProcName = "get2HPrice"
   'On Error GoTo errProc
   
   get2HPrice = 0#
   
   lsSQL = "SELECT b.sClassIDx" & _
               ", IFNULL(c.nMinRatex, 0) nMinRatex" & _
               ", IFNULL(d.nSelPrice, 0) nSelPrice" & _
            " FROM MC_Serial a" & _
               " LEFT JOIN MC_Model e" & _
                  " ON a.sModelIDx = e.sModelIDx" & _
               " LEFT JOIN MC_Serial_Classification b" & _
                  " ON a.sSerialID = b.sSerialID" & _
               " LEFT JOIN MC_2H_Classification c" & _
                  " ON b.sClassIDx = c.sClassIDx" & _
                     " AND c.cHotModel = e.cHotItemx" & _
               " LEFT JOIN MC_Model_Price d" & _
                  " ON a.sModelIDx = d.sModelIDx" & _
            " WHERE a.sSerialID = " & strParm(lsSerialID)
   Set loRS = New Recordset
   Debug.Print lsSQL
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then GoTo endProc
   
   'get configuration value for minimum discount deduction
   lnMinRatex = CDbl(IFNull(p_oAppDrivr.getConfiguration("2HRateAllw"), 0))
   
   lnMinRatex = loRS("nMinRatex") - lnMinRatex '2%
   
   lnMinRatex = loRS("nSelPrice") - (loRS("nSelPrice") * lnMinRatex / 100)
   
   get2HPrice = Round(lnMinRatex / 50) * 50
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsSerialID & " )"
End Function

'mac 2023.11.29
Private Function isBNOS(ByVal fsSerialID As String) As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   
   lsSQL = "SELECT b.nBNOSSRPx" & _
               " FROM MC_BNOS_Classify_Master a" & _
                  ", MC_BNOS_Classify_Detail b" & _
               " WHERE a.sTransNox = b.sTransNox" & _
                  " AND b.sSerialID = " & strParm(fsSerialID) & _
                  " AND a.cTranStat = '2'" & _
               " ORDER BY a.dApproved DESC LIMIT 1"

   Set loRS = New Recordset
   
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   isBNOS = Not loRS.EOF
End Function
