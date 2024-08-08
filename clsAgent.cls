VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Sales Agent Object
'
' Copyright 2015 and Beyond
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
'  XerSys [ 05/09/2015 09:50 pm ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsAgent"

Private p_oAppDrivr As clsAppDriver
Private p_oClient As clsNeoClient
Private p_oMaster As Recordset

Private p_sBranchCd As String
Private p_sSQLMaster As String

Private p_sFullName As String
Private p_sAddressx As String
Private p_sSpouseNm As String
Private p_sAgentSQL As String

Private pbInitTran As Boolean
Private pbRecExist As Boolean

Private pbModified As Boolean
Private pnCtr As Integer
Private pbShwMsg As Boolean

Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal value As String)
   p_sBranchCd = value
End Property

Property Get Master(Index As Variant) As Variant
   'on error Resume Next

   If Not pbInitTran Then
      Call showNotInitMsg
      Exit Property
   End If
   
   If p_oMaster.EOF Then
'      MsgBox "Invalid Row Value Detected!!!", vbCritical, "Warning"
      Exit Property
   End If
   
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, 1, 2, "nentrynox", "sbarrcode", "sdescript"
      Master = p_oSPOrder.Master(Row, Index)
   Case 3, 4, 7, 8, 9, 10, 11, 12, 13, _
         "nquantity", "nrecorder", _
         "ncancelld", "nqtyonhnd", "nresvordr", _
         "nbackordr", "nfloatqty", "navemonsl", _
         "cclassify"
      If IsNumeric(Index) Then Index = Index + 2
      
      Master = p_oSPOrder.Master(Row, Index)
   Case 5, "nissuedxx"
      Master = p_oMaster("nIssuedxx")
   Case 6, "norderqty"
      Master = p_oSPOrder.Master(Row, Index) - p_oMaster("nReceived")
   Case 14, "xissuedxx"
      Master = p_oSPOrder.Master(Row, "nIssuedxx")
   Case 15, "xresvordr"
      Master = p_oMaster("nResvOrdr")
   Case 16, "xbackordr"
      Master = p_oMaster("nBackOrdr")
   Case 17, "xfloatqty"
      Master = p_oMaster("nFloatQty")
   Case 18, "xqtyonhnd"
      Master = p_oMaster("nQtyOnHnd")
   Case 19, "smodelnme"
      Master = p_oMaster("sModelNme")
   Case 20, "spartsidx"
      Master = p_oSPOrder.Master(Row, "sPartsIDx")
   Case 21, "xbarrcode"
      Master = p_oMaster("xBarrCode")
   Case 22, "xpartsidx"
      Master = p_oMaster("xPartsIDx")
   Case Else
'      MsgBox "Invalid property index detected!", vbCritical, "Warning"
      Exit Property
   End Select
End Property

Property Let Master(ByVal Row As Long, Index As Variant, value As Variant)
   'on error Resume Next

   If pbInitTran = False Then
      Call showNotInitMsg
      Exit Property
   End If

   If Row > p_oMaster.RecordCount - 1 Then
      MsgBox "Invalid Row Value Detected!!!", vbCritical, "Warning"
      Exit Property
   End If
         
   p_oMaster.Move Row, adBookmarkFirst
   If Not IsNumeric(Index) Then Index = LCase(Index)
   Select Case Index
   Case 0, 1, 2, 3, 4, "nentrynox", "sbarrcode", "sdescript", "nquantity", "nrecorder"
   Case 5, "nissuedxx"
      If Row > p_oMaster.RecordCount - 1 Then
         MsgBox "Invalid Row Value Detected!!!", vbCritical, "Warning"
         Exit Property
      End If
      
      If Not IsNumeric(value) Then
         MsgBox "Invalid Property Value Detected!!!", vbCritical, "Warning"
         Exit Property
      End If
      
      If CLng(value) > p_oMaster("nQtyOnHnd") Then
         MsgBox "Issue quantity is greater than actual quantity on hand!", vbCritical, "Warning"
         Exit Property
      End If
      
      p_oMaster.Move Row, adBookmarkFirst
      p_oMaster("nIssuedxx") = value
   Case 6, 7, 8, 9, 10, 11, "ncancelld", "norderqty", "nqtyonhnd", "nresvordr", "nbackordr", "nfloatqty"
   Case 12, 13, "navemonsl", "cclassify"
   Case 14, "xissuedxx"
   Case 15, 16, 17, 18, 19, "xresvordr", "xbackordr", "xfloatqty", "xqtyonhnd", "smodelnme"
   Case 20, "spartsidx"
   Case Else
      MsgBox "Invalid property index detected!", vbCritical, "Warning"
      Exit Property
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If Not pbInitTran Then
      Call showNotInitMsg
      Exit Property
   End If
   
   Master = p_oSPOrder.Master(Index)
End Property

Property Let Master(Index As Variant, value As Variant)
End Property

Property Let TransStatus(ByVal value As Integer)
   p_nTranStat = value
End Property

Property Get ItemCount() As Long
   If Not pbInitTran Then
      Call showNotInitMsg
      Exit Property
   End If

   ItemCount = p_oMaster.RecordCount
End Property

Property Let Parent(ByVal value As String)
   p_sParent = value
End Property

Function InitTransaction(Optional oAppDriver As Variant) As Boolean
   Dim lsProcName As String
   Dim lsCondition As String
   
   lsProcName = "InitTransaction"
   On Error GoTo errProc

   If pbInitTran = False Then
      If IsMissing(oAppDriver) Then
         If isAppDriverOK(p_oAppDrivr) = False Then
            Debug.Print "Uninitialized application driver!"
            GoTo endProc
         End If
      Else
         If TypeName(oAppDriver) <> "clsAppDriver" Then
            Debug.Print "Invalid application driver detected!"
            GoTo endProc
         End If
         Set p_oAppDrivr = oAppDriver
      End If
      If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
      pbInitTran = True
   End If
   
   p_sSQLMaster = "SELECT a.nEntryNox" & _
                     ", a.sPartsIDx" & _
                     ", a.nIssuedxx" & _
                     ", IFNULL(b.nResvOrdr, 0) nResvOrdr" & _
                     ", IFNULL(b.nBackOrdr, 0) nBackOrdr" & _
                     ", IFNULL(b.nFloatQty, 0) nFloatQty" & _
                     ", IFNULL(b.nQtyOnHnd, 0) nQtyOnHnd" & _
                     ", d.sModelNme" & _
                     ", c.nSelPrice" & _
                     ", c.sBarrCode xBarrCode" & _
                     ", c.sPartsIDx xPartsIDx" & _
                     ", h.nReceived"
   p_sSQLMaster = p_sSQLMaster & _
                  " FROM SP_Stock_Request_Master a" & _
                     " LEFT JOIN SP_Supersession e ON a.sPartsIDx = e.sPartsIDx" & _
                     " LEFT JOIN SP_Supersession g ON e.sSuperIDx = g.sSuperIDx" & _
                     " LEFT JOIN SP_Inventory b ON g.sPartsIDx = b.sPartsIDx AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                     " LEFT JOIN Spareparts c ON g.sPartsIDx = c.sPartsIDx" & _
                     " LEFT JOIN SP_Model d ON c.sModelIDx = d.sModelIDx" & _
                     " LEFT JOIN (" & _
                        " SELECT z.sPartsIDx" & _
                           ", z.sReplacID" & _
                           ", z.nQuantity nReceived" & _
                           ", y.sTransNox" & _
                           ", z.sOrderNox" & _
                        " FROM SP_PO_Receiving_Master y" & _
                           ", SP_PO_Receiving_Master z" & _
                        " WHERE y.sTransNox = z.sTransNox" & _
                           " AND y.cTranStat <> " & strParm(xeStateCancelled) & _
                           " AND z.sOrderNox = " & strParm("«empty»") & _
                        " ) h" & _
                        " ON a.sTransNox = h.sOrderNox" & _
                           " AND ( a.sPartsIDx = h.sPartsIDx" & _
                              " OR a.sPartsIDx = h.sReplacID)" & _
                  " WHERE a.nQuantity - a.nCancelld > a.nIssuedxx + IF(a.nOrderQty > a.nQuantity, a.nQuantity, a.nOrderQty) - IFNULL(h.nReceived, 0)" & _
                     " AND a.sTransNox = " & strParm("«empty»") & _
                  " ORDER BY nEntryNox"

   Set p_oMaster = New Recordset
   Debug.Print p_sSQLMaster
   p_oMaster.Open p_sSQLMaster, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Debug.Print "Record Count:", p_oMaster.RecordCount
      
   Set p_oSPOrder = New clsSPStockOrder
   With p_oSPOrder
      Set .AppDriver = p_oAppDrivr
      .Parent = pxeMODULENAME
      .DisplayMessage = False
      .Distribute = True
   
      If Not .InitTransaction Then GoTo endProc
   End With
   
   Call InitPeriod
   InitTransaction = True
   pbInitTran = InitTransaction

endProc:
   Exit Function
errProc:
  ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lbTransfer As Boolean
   Dim lsSupplier As String

   lsProcName = "SaveTransaction"
   On Error GoTo errProc

   If Not pbInitTran Then
      Call showNotInitMsg
      GoTo endProc
   End If
   
   If Not pbRecExist Then
      Call showNoRecMsg
      GoTo endProc
   End If
   
   p_oAppDrivr.BeginTrans
   ' open SP Order for the updates
   With p_oSPOrder
      ' update first the order of the branch
      p_oMaster.MoveFirst
      Do While Not p_oMaster.EOF
         ' check for stock to be transferred
         If p_oMaster("nIssuedxx") > 0 Then lbTransfer = True
         
         lnCtr = p_oMaster("nEntryNox")
         .Master(lnCtr - 1, "nIssuedxx") = .Master(lnCtr - 1, "nIssuedxx") + p_oMaster("nIssuedxx")
         
         p_oMaster.MoveNext
      Loop
   
      If lbTransfer Then
         .Destination = Left(.Master("sTransNox"), pxeBranchCodeLen)
         
         If Not p_oSPOrder.SaveTransaction() Then GoTo endWithRoll
         If Not saveTransfer() Then GoTo endWithRoll
      End If
      p_oAppDrivr.CommitTrans
      
   End With
   
   SaveTransaction = True

endProc:
  Exit Function
endWithRoll:
   p_oAppDrivr.RollbackTrans
   GoTo endProc
errProc:
  ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(sTransNo As String) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "OpenTransaction"
   On Error GoTo errProc

   If Not pbInitTran Then
      Call showNotInitMsg
      GoTo endProc
   End If
   
   lsSQL = Replace(p_sSQLMaster, "«empty»", sTransNo)
'   lsSQL = AddCondition(lsSQL, "a.sTransNox = " & strParm(sTransNo))
   
   Debug.Print lsSQL
   Set p_oMaster = New Recordset
   p_oMaster.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oMaster.ActiveConnection = Nothing
   If Not p_oMaster.EOF Then
      If Not p_oSPOrder.OpenTransaction(sTransNo) Then GoTo endProc
      pbRecExist = True
   End If
   
   Do Until p_oMaster.EOF
      p_oMaster("nIssuedxx") = 0
      
      If IsNull(p_oMaster("nReceived")) Then
         p_oMaster("nReceived") = 0
      End If
      p_oMaster.MoveNext
   Loop
   
   OpenTransaction = pbRecExist
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sTransNo & " )"
End Function

'Function SearchTransaction(Optional sSearch As Variant) As Boolean
'   Dim lsProcName As String
'   Dim lsSQL As String
'   Dim lasSelect() As String
'   Dim lors As Recordset
'
'   lsProcName = "SearchTransaction"
'   On Error Goto errProc
'   SearchTransaction = False
'
'   If pbInitTran = False Then GoTo endProc
'
'   lsSQL = "SELECT DISTINCT a.sTransNox" & _
'               ", b.sBranchNm" & _
'               ", a.dTransact" & _
'            " FROM SP_Stock_Request_Master a" & _
'                  " LEFT JOIN SP_PO_Master d" & _
'                     " LEFT JOIN SP_PO_Master e" & _
'                        " ON d.sTransNox = e.sTransNox" & _
'                     " ON a.sTransNox = d.sOrderNox" & _
'                        " AND ( d.sBranchCd = " & strParm(p_sBranchCd) & _
'                           IIf(p_sBranchCd <> pxeMainWH, "", _
'                           " OR d.sBranchCd IN " & strParm(pxeMainRc)) & " )" & _
'               ", Branch b" & _
'               ", SP_Stock_Request_Master c" & _
'            " WHERE a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
'               " AND a.sTransNox = c.sTransNox" & _
'               " AND ( a.cTranStat = " & strParm(xeStateClosed) & _
'                  " OR a.cTranStat = " & strParm(xeStatePosted) & " )" & _
'               " AND c.nQuantity - c.nCancelld > c.nIssuedxx + c.nOrderQty - IFNULL(e.nReceived, 0)" & _
'               IIf(p_sReqBranchCd <> "", " AND b.sBranchCd = " & strParm(p_sReqBranchCd), "")
'
'   If Not IsMissing(sSearch) Then
'      If p_oSPOrder.Master("sTransNox") = sSearch Then GoTo endProc
'      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(sSearch & "%"))
'   End If
'
'   Debug.Print lsSQL
'   Set lors = New Recordset
'   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
'
'   If lors.EOF Then GoTo endProc
'
'   If lors.RecordCount > 1 Then
'      lsSQL = KwikBrowse(p_oAppDrivr, lors, "sBranchCd»sBranchNm»dTransact", _
'                        "Transaction #»Branch»Date", "@@@@»@»MMM DD, YYYY")
'      If lsSQL = "" Then GoTo endProc
'
'      lasSelect = Split(lsSQL, "»")
'
'      Call lors.Find("sTransNox = " & strParm(lasSelect(0)), 0, adSearchForward)
'   End If
'
'   SearchTransaction = OpenTransaction(lors("sTransNox"))
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & sSearch & " )"
'End Function

Function SearchTransaction(Optional sSearch As Variant) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lasSelect() As String
   Dim lors As Recordset

   lsProcName = "SearchTransaction"
   On Error GoTo errProc
   SearchTransaction = False

   If pbInitTran = False Then GoTo endProc

   lsSQL = "SELECT DISTINCT a.sTransNox" & _
               ", b.sBranchNm" & _
               ", a.dTransact" & _
            " FROM SP_Stock_Request_Master a" & _
                  " LEFT JOIN SP_PO_Master d" & _
                     " LEFT JOIN SP_PO_Master e" & _
                        " ON d.sTransNox = e.sTransNox" & _
                     " ON a.sTransNox = d.sOrderNox" & _
                        " AND ( d.sBranchCd = " & strParm(p_sBranchCd) & _
                           IIf(p_sBranchCd <> pxeMainWH, "", _
                           " OR d.sBranchCd IN " & pxeMainRc) & " )" & _
               ", Branch b" & _
               ", SP_Stock_Request_Master c" & _
            " WHERE a.sTransNox LIKE CONCAT(b.sBranchCd, '%')" & _
               " AND a.sTransNox = c.sTransNox" & _
               " AND ( a.cTranStat = " & strParm(xeStateClosed) & _
                  " OR a.cTranStat = " & strParm(xeStatePosted) & " )" & _
               " AND c.nQuantity - c.nCancelld > c.nIssuedxx + c.nOrderQty - IFNULL(e.nReceived, 0)" & _
               IIf(p_sReqBranchCd <> "", " AND b.sBranchCd = " & strParm(p_sReqBranchCd), "")
               
   If Not IsMissing(sSearch) Then
      If p_oSPOrder.Master("sTransNox") = sSearch Then GoTo endProc
      lsSQL = AddCondition(lsSQL, "a.sTransNox LIKE " & strParm(sSearch & "%"))
   End If
   
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If lors.EOF Then GoTo endProc
    
   If lors.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, lors, "sBranchCd»sBranchNm»dTransact", _
                        "Transaction #»Branch»Date", "@@@@»@»MMM DD, YYYY")
      If lsSQL = "" Then GoTo endProc
      
      lasSelect = Split(lsSQL, "»")
      
      Call lors.Find("sTransNox = " & strParm(lasSelect(0)), 0, adSearchForward)
   End If
   
   SearchTransaction = OpenTransaction(lors("sTransNox"))

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sSearch & " )"
End Function

Function SearchBranch(ByVal lsValue As String, ByVal lbSearch As Boolean) As String
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getBranch"
   On Error GoTo errProc
   
   SearchBranch = p_sReqBranchNm
   
   If lsValue <> "" Then
      If lsValue = p_sReqBranchNm Then GoTo endProc
      
      If lbSearch Then
         lsSQL = "a.sBranchNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSQL = "a.sBranchNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endProc
   End If
   
   lsSQL = "SELECT a.sBranchCd" & _
               ", a.sBranchNm" & _
               ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName) xAddressx" & _
            " FROM Branch a" & _
               ", TownCity b" & _
               ", Province c" & _
            " WHERE a.sTownIDxx = b.sTownIDxx" & _
               " AND b.sProvIDxx = c.sProvIDxx" & _
               " AND a.sBranchCd <> " & strParm(p_sBranchCd) & _
               " AND sBranchCd LIKE " & strParm("M%") & _
               IIf(lsSQL = "", "", " AND " & lsSQL)
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If lors.EOF Then
      If Not lbSearch Then
         p_sReqBranchCd = ""
         p_sReqBranchNm = ""
         p_sAddressx = ""
      End If
      GoTo endProc
   End If
   
   If lors.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, lors _
                           , "sBranchCd»sBranchNm" _
                           , "ID»Branch")
      If lsSQL = "" Then
         GoTo endProc
      Else
         lasMaster = Split(lsSQL, "»")
         
         lors.MoveFirst
         lors.Find "sBranchCd = " & strParm(lasMaster(0)), 0, adSearchForward
      End If
   End If
      
   p_sReqBranchCd = lors("sBranchCd")
   p_sReqBranchNm = lors("sBranchNm")
   p_sAddressx = lors("xAddressx")
   
   SearchBranch = p_sReqBranchNm
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lbSearch & " )"
End Function

'Function GetBranchOrder() As Recordset
'   Dim lors As Recordset
'   Dim lsSQL As String
'   Dim lsProcName As String
'
'   lsProcName = "GetBranchOrder"
'   On Error Goto errProc
'
'   If p_sReqBranchCd = "" Then
'      Set GetBranchOrder = Nothing
'      GoTo endProc
'   End If
'
'   lsSQL = "SELECT a.sTransNox" & _
'               ", a.dTransact" & _
'               ", COUNT(b.sPartsIDx) xUnserved" & _
'            " FROM SP_Stock_Request_Master a" & _
'               ", SP_Stock_Request_Master b" & _
'            " WHERE a.sTransNox = b.sTransNox" & _
'               " AND ( a.cTranStat = " & strParm(xeStateClosed) & _
'                  " OR a.cTranStat = " & strParm(xeStatePosted) & " )" & _
'               " AND a.sTransNox LIKE " & strParm(p_sReqBranchCd & "%") & _
'               " AND b.nQuantity - b.nCancelld > b.nIssuedxx + b.nOrderQty" & _
'            " GROUP BY a.sTransNox"
'   lsSQL = lsSQL & " UNION ALL" & _
'            " SELECT a.sTransNox" & _
'               ", a.dTransact" & _
'               ", COUNT(b.sPartsIDx) xUnserved" & _
'            " FROM SP_Stock_Request_Master a" & _
'               ", SP_Stock_Request_Master b" & _
'               ", SP_PO_Master c" & _
'               ", SP_PO_Master d" & _
'            " WHERE a.sTransNox = b.sTransNox" & _
'               " AND a.sTransNox = c.sOrderNox" & _
'               " AND c.sTransNox = d.sTransNox AND b.sPartsIDx = d.sPartsIDx" & _
'               " AND ( c.sBranchCd = " & strParm(p_sBranchCd) & _
'                        IIf(p_sBranchCd <> pxeMainWH, "", _
'                        " OR c.sBranchCd IN " & pxeMainRc) & " )" & _
'               " AND ( a.cTranStat = " & strParm(xeStateClosed) & _
'                  " OR a.cTranStat = " & strParm(xeStatePosted) & " )" & _
'               " AND a.sTransNox LIKE " & strParm(p_sReqBranchCd & "%") & _
'               " AND b.nQuantity - b.nCancelld > b.nIssuedxx + b.nOrderQty - d.nQuantity" & _
'            " GROUP BY a.sTransNox"
'
'   Set lors = New Recordset
'   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
'   Debug.Print lsSQL
'   Set GetBranchOrder = lors
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & " )"
'End Function
'

Function GetBranchOrder() As Recordset
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   
   lsProcName = "GetBranchOrder"
   On Error GoTo errProc
   
   If p_sReqBranchCd = "" Then
      Set GetBranchOrder = Nothing
      GoTo endProc
   End If
   
   lsSQL = "SELECT a.sTransNox" & _
               ", a.dTransact" & _
               ", COUNT(b.sPartsIDx) xUnserved" & _
            " FROM SP_Stock_Request_Master a" & _
               ", SP_Stock_Request_Master b" & _
            " WHERE a.sTransNox = b.sTransNox" & _
               " AND ( a.cTranStat = " & strParm(xeStateClosed) & _
                  " OR a.cTranStat = " & strParm(xeStatePosted) & " )" & _
               " AND a.sTransNox LIKE " & strParm(p_sReqBranchCd & "%") & _
               " AND b.nQuantity - b.nCancelld > b.nIssuedxx + IF(b.nOrderQty > b.nQuantity, b.nQuantity, b.nOrderQty)" & _
               IIf(psOrderType = "1", " AND b.nResvOrdr > 0", "") & _
               IIf(psOrderType = "2", " AND b.nResvOrdr = 0", "") & _
            " GROUP BY a.sTransNox"
   lsSQL = lsSQL & " UNION " & _
            " SELECT a.sTransNox" & _
               ", a.dTransact" & _
               ", b.nQuantity - (b.nCancelld + b.nIssuedxx) xUnserved" & _
            " FROM SP_Stock_Request_Master a" & _
               ", SP_Stock_Request_Master b" & _
               ", SP_PO_Receiving_Master c" & _
            " WHERE a.sTransNox = b.sTransNox" & _
               " AND a.sTransNox = c.sOrderNox" & _
               " AND ( b.sPartsIDx = c.sPartsIDx" & _
                  " OR b.sPartsIDx = c.sReplacID )" & _
               " AND a.sTransNox LIKE " & strParm(p_sReqBranchCd & "%") & _
               " AND b.nQuantity - (b.nCancelld + b.nIssuedxx) >= IF(c.nQuantity > b.nOrderQty, b.nOrderQty, c.nQuantity)" & _
               IIf(psOrderType = "1", "AND b.nResvOrdr > 0", "") & _
               IIf(psOrderType = "2", "AND b.nResvOrdr = 0", "") & _
            " GROUP BY a.sTransNox" & _
            " HAVING xUnserved > 0" & _
            " ORDER BY sTransNox"
            
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Debug.Print lsSQL
   Set GetBranchOrder = lors
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function saveTransfer() As Boolean
   Dim lsProcName As String
   Dim loFormTransfer As frmSP_TransferStock
   Dim loTransfer As clsSPTransfer
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lnTranTotal As Double
   Dim loReport As CRAXDRT.Report
   Dim loRepApp As New CRAXDRT.Application
   
   lsProcName = "saveTransfer"
   On Error GoTo errProc
   
   ' for printing purposes
   Set lors = New ADODB.Recordset
   
   lors.Fields.Append "nEntryNo", adInteger, 3
   lors.Fields.Append "sBarrCode", adVarChar, 23
   lors.Fields.Append "sDescription", adVarChar, 60
   lors.Fields.Append "sModel", adVarChar, 30
   lors.Fields.Append "nQuantity", adInteger, 5
   lors.Fields.Append "nUnitPrice", adDouble, 20
   lors.Fields.Append "nTotal", adDouble, 20
   lors.Open

   Set loTransfer = New clsSPTransfer
   With loTransfer
      Set .AppDriver = p_oAppDrivr
      .SaveOrder = False
      .DisplayConfirmation = False
      If Not .InitTransaction() Then GoTo endProc
      If Not .NewTransaction() Then GoTo endProc
      
      Set loFormTransfer = New frmSP_TransferStock
      Set loFormTransfer.AppDriver = p_oAppDrivr
      Call Load(loFormTransfer)
      
      lnCtr = -1
      p_oMaster.MoveFirst
      Do Until p_oMaster.EOF
         If p_oMaster("nIssuedxx") > 0 Then
            If lnRow > 0 Then .AddMaster
            
            .Master(lnRow, "sPartsIDx") = p_oMaster("xPartsIDx")
            .Master(lnRow, "nQuantity") = p_oMaster("nIssuedxx")
            .Master(lnRow, "nUnitPrce") = p_oMaster("nSelPrice")
            ' XerSys - 2014-10-08
            '  Assign the superseeded order to
            If p_oMaster("xPartsIDx") <> p_oMaster("sPartsIDx") Then
               .Master(lnRow, "sOrigIDxx") = p_oMaster("sPartsIDx")
               Debug.Print .Master(lnRow, "sOrigIDxx")
            End If
            
            With loFormTransfer.MSFlexGrid1
               .Rows = lnRow + 2
               lnCtr = p_oMaster("nEntryNox") - 1
               .TextMatrix(lnRow + 1, 0) = lnRow + 1
'               .TextMatrix(lnRow + 1, 1) = p_oSPOrder.Master(lnCtr, "sBarrCode")
               .TextMatrix(lnRow + 1, 1) = p_oMaster("xBarrCode")
               .TextMatrix(lnRow + 1, 2) = p_oSPOrder.Master(lnCtr, "sDescript")
               .TextMatrix(lnRow + 1, 3) = IFNull(p_oMaster("sModelNme"), "")
               .TextMatrix(lnRow + 1, 4) = IFNull(p_oMaster("nQtyOnHnd"), 0)
               .TextMatrix(lnRow + 1, 5) = p_oMaster("nIssuedxx")
               .TextMatrix(lnRow + 1, 6) = Format(p_oMaster("nSelPrice"), "#,##0.00")
               
               lors.AddNew
               lors("nEntryNo").value = .TextMatrix(lnRow + 1, 0)
               lors("sBarrCode").value = .TextMatrix(lnRow + 1, 1)
               lors("sDescription").value = .TextMatrix(lnRow + 1, 2)
               lors("sModel").value = .TextMatrix(lnRow + 1, 3)
               lors("nQuantity").value = .TextMatrix(lnRow + 1, 5)
               lors("nUnitPrice").value = .TextMatrix(lnRow + 1, 6)
               lors("nTotal").value = Round(lors("nQuantity") * lors("nUnitPrice"), 2)
               lnTranTotal = lnTranTotal + lors("nTotal")
            End With
            
            lnRow = lnRow + 1
         End If
         
         p_oMaster.MoveNext
      Loop
      
      If lnCtr >= 0 Then
         .Master("sDestinat") = p_sReqBranchNm
         .Master("dTransact") = p_oAppDrivr.ServerDate
         
         loFormTransfer.TransactionNo = Format(loTransfer.Master("sTransNox"), "@@@@-@@-@@@@@@")
         loFormTransfer.Branch = p_sReqBranchNm
         loFormTransfer.Address = p_sAddressx
         loFormTransfer.Show 1
         
         If loFormTransfer.Cancelled Then
            Unload loFormTransfer
            GoTo endProc
         End If
         
         .Master("dTransact") = loFormTransfer.TransactionDate
         .Master("nDiscount") = loFormTransfer.Discount
         .Master("sRemarksx") = loFormTransfer.Remarks
         .Master("sOrderNox") = p_oSPOrder.Master("sTransNox")
         
         If Not .SaveTransaction() Then GoTo endProc
         
         On Error Resume Next
         If MsgBox("Would you like to Print Transfer Now?", vbInformation + vbYesNo, "Confirm") = vbYes Then
            ' proceed with printing
            
            ' assign important info to the report
            Set loReport = loRepApp.OpenReport(p_oAppDrivr.AppPath & "\Reports\Branch Transfer.rpt")
            loReport.DiscardSavedData
            loReport.FieldMappingType = crAutoFieldMapping
            loReport.Database.SetDataSource lors
            
            loReport.Sections("RHa").ReportObjects("txtRefNo").SetText "SP" & "-" & Right(.Master("sTransNox"), 8)
            loReport.Sections("RHa").ReportObjects("txtDate").SetText loFormTransfer.txtField(3).Text
            loReport.Sections("RHb").ReportObjects("txtToBranch").SetText loFormTransfer.txtField(1)
            loReport.Sections("RHb").ReportObjects("txtToAddress").SetText loFormTransfer.txtField(2).Text
            loReport.Sections("RF").ReportObjects("txtDisc").SetText loFormTransfer.txtField(4).Text
            loReport.Sections("RF").ReportObjects("txtRemarks").SetText loFormTransfer.txtField(5).Text
            loReport.Sections("RF").ReportObjects("txtNet").SetText Format(Round(lnTranTotal * (100 - .Master("nDiscount")) / 100, 2), "#,##0.00")
            loReport.Sections("RF").ReportObjects("txtApproved").SetText loFormTransfer.txtField(6).Text
            loReport.Sections("RF").ReportObjects("txtPrepared").SetText p_oAppDrivr.UserName
            
            Do While True
               loReport.PrintOutEx False, 1
               
               If MsgBox("Would you like to Re-Print Transfer?", vbInformation + vbYesNo) <> vbYes Then
                  Exit Do
               End If
            Loop
         End If
         On Error GoTo errProc
      Else
         MsgBox "No Modified Record Detected!", vbCritical, "Warning"
         Call Unload(loFormTransfer)
         GoTo endProc
      End If
   End With
   
   saveTransfer = True
   
endProc:
   Unload loFormTransfer
         
   Set lors = Nothing
   Set loReport = Nothing
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

Private Sub showNotInitMsg()
   MsgBox "Transaction is not yet Initialized!" & vbCrLf & _
            "Please verify your entry then Try Again!", vbCritical, "Warning"
End Sub

Private Sub showNoRecMsg()
   MsgBox "No transaction is currently loaded!" & vbCrLf & _
            "Please verify your entry then Try Again!", vbCritical, "Warning"
End Sub

Private Sub InitPeriod()
10       Dim lsProcName As String
20       Dim ldDate As Date
30       Dim lnCtr As Integer
40       Dim loSPOther As clsSPOthers
   
50       lsProcName = "getPeriod"
   On Error GoTo errProc
      
60       Set loSPOther = New clsSPOthers
70       With loSPOther
80          Set .AppDriver = p_oAppDrivr
90          Call .InitTransaction
100      End With
   
110      ReDim p_asPeriod(loSPOther.Master("nNoMonths"))
120      ldDate = Date
130      For lnCtr = 0 To loSPOther.Master("nNoMonths")
140         p_asPeriod(lnCtr) = Format(ldDate, "yyyymm")
150         ldDate = DateAdd("m", -1, ldDate)
160      Next
   
endProc:
170      Exit Sub
errProc:
180      ShowError lsProcName & "( " & " )"
End Sub

Function GetHistory(ByVal PartID As String) As Variant
10       Dim lsProcName As String
20       Dim lsSQL As String
30       Dim lors As Recordset, loTmp As Recordset
40       Dim lasHistory() As Variant
50       Dim lnCtr As Integer
   
60       lsProcName = "GetHistory"
   On Error GoTo errProc
   
70       ReDim lasHistory(UBound(p_asPeriod), 2)
80       Set lors = New Recordset
90       With lors
100         For lnCtr = 0 To UBound(p_asPeriod)
110            lasHistory(lnCtr, 0) = p_asPeriod(lnCtr)

120            lsSQL = "SELECT SUM(nQtyInxxx) nOrderQty" & _
                  " FROM SP_Inventory_Ledger" & _
                  " WHERE sBranchCd = " & strParm(p_sBranchCd) & _
                     " AND sPartsIDx = " & strParm(PartID) & _
                     " AND DATE_FORMAT(dTransact, '%Y%m') = " & strParm(p_asPeriod(lnCtr)) & _
                     " AND sSourceCd IN (" & strParm(pxeSPAcceptDelivery) & _
                        ", " & strParm(pxeSPPurchaseReceiving) & _
                        ", " & strParm(pxeSPAcceptWarrantyTransfer) & ")"

130            .Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
140            Debug.Print lsSQL
150            If .EOF Then
160               lasHistory(lnCtr, 2) = 0
170            Else
180               lasHistory(lnCtr, 2) = IFNull(.Fields("nOrderQty"), 0)
190            End If
200            .Close
210         Next
      
220         lsSQL = "SELECT DATE_FORMAT(dModified, '%Y%m') sPeriodxx" & _
                  ", COUNT(sTransNox) nInquiryx" & _
               " FROM SP_StockInquiry_Master" & _
               " WHERE sTransNox LIKE " & strParm(p_sBranchCd & "%") & _
                  " AND sPartsIDx = " & strParm(PartID) & _
                  " AND DATE_FORMAT(dModified, '%Y%m') BETWEEN " & _
                     strParm(p_asPeriod(UBound(p_asPeriod))) & _
                     " AND " & strParm(p_asPeriod(0)) & _
               " GROUP BY sPeriodxx" & _
               " ORDER BY sPeriodxx DESC"
               
230         .Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
240         Set .ActiveConnection = Nothing
      
250         lnCtr = 0
260         Debug.Print lsSQL
270         If .EOF Then
280            For lnCtr = 0 To UBound(p_asPeriod)
290               lasHistory(lnCtr, 1) = 0
300            Next
310         Else
320            .MoveFirst
         
330            Do Until .EOF
340               If .Fields("sPeriodxx") = p_asPeriod(lnCtr) Then
350                  lasHistory(lnCtr, 1) = .Fields("nInquiryx")
360                  .MoveNext
370               Else
380                  lasHistory(lnCtr, 1) = 0
390               End If
400               lnCtr = lnCtr + 1
410            Loop
420         End If
430      End With
   
440      GetHistory = lasHistory
   
endProc:
450      Exit Function
errProc:
460      ShowError lsProcName & "(" & ")"
End Function


