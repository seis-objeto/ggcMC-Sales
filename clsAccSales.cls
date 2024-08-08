VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccSales"
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
'     Motorcycle Sales Accessories Class Object
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
'  XerSys [ 12/15/2004 02:26 pm ]
'     Start creating this object.
'  XerSys [ 02/21/2005 05:25 pm ]
'     Adjust this object to fit to the new implementation of the application driver
'  Kalyptus [ 12/11/2006 03:28 pm]
'     Implements the BugTracker
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private WithEvents p_oBaseClas As clsMasterDetail
Attribute p_oBaseClas.VB_VarHelpID = -1

Private p_sBranchCd As String

Private pbInitTran As Boolean
Private pbModified As Boolean
Private pnEditMode As xeEditMode

Private pnCtr As Long

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
   Case 1, "sdescript"
      getDetail Row, Value, False
   Case Else
      p_oBaseClas.Detail(Row, Index) = Value
   End Select
End Property

Property Get Master(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   
   Master = p_oBaseClas.Master(Index)
End Property

Property Let Master(Index As Variant, Value As Variant)
   If pbInitTran = False Then Exit Property
   
   p_oBaseClas.Master(Index) = Value
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oBaseClas.ItemCount
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
      .DetailTable = "MC_SO_Accessories"

      .MasterQuery = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.dTransact" & _
                        ", a.sDRNoxxxx" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                        ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName) AS xAddressx" & _
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
                        ", b.sDescript" & _
                        ", c.nQtyOnHnd" & _
                        ", a.nQuantity" & _
                        ", a.nGivenxxx" & _
                        ", a.sTransNox" & _
                        ", a.sAcsModID" & _
                        ", c.nQtyOnHnd + a.nGivenxxx xQtyOnHnd" & _
                     " FROM " & .DetailTable & " a" & _
                        ", Accessories b" & _
                        ", Access_Model c" & _
                     " WHERE a.sAcsModID = c.sAcsModID" & _
                        " AND b.sAccessID = c.sAccessID" & _
                        " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                     " ORDER BY a.nEntryNox"
      .Reference = "a.sTransNox|a.sTransNox"
      
      InitTransaction = .InitTransaction

      .BrowseQuery(0) = "SELECT" & _
                           "  a.sTransNox" & _
                           ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) AS xFullName" & _
                           ", b.sCompnyNm" & _
                           ", a.dTransact" & _
                        " FROM " & .MasterTable & " a" & _
                           ", Client_Master b" & _
                        " WHERE a.sClientID = b.sClientID" & _
                        " ORDER BY b.sCompnyNm" & _
                           ", a.dTransact"
      .BrowseColumn(0) = "sTransNox»xFullName»sCompnyNm»dTransact"
      .BrowseCriteria(0) = "a.sTransNox»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)»b.sCompnyNm»a.dTransact"
      .BrowseTitle(0) = "Transaction No»Customer Name»Company Name»Date"
      .BrowseFormat(0) = "@@-@@@@@@»@»@»MMMM DD, YYYY"
      
      .BrowseDetailQuery(0) = "SELECT" & _
                                 "  a.sAcsModID" & _
                                 ", b.sDescript" & _
                                 ", a.nQtyOnHnd" & _
                              " FROM Access_Model a" & _
                                 ", Accessories b" & _
                              " WHERE a.sAccessID = b.sAccessID" & _
                                 " AND a.sBranchCd = " & strParm(p_sBranchCd) & _
                                 " AND a.cRecdStat = " & strParm(xeRecStateActive)
      .BrowseDetailColumn(0) = "sAcsModID»sDescript»nQtyOnHnd"
      .BrowseDetailTitle(0) = "Bar Code»Description»QOH"
   End With
   pbInitTran = InitTransaction
   
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

Function SearchTransaction(Optional sSearch As Variant, Optional bByCode As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsValue As String
   Dim lsField As String
   Dim lsDescript As String
   
   lsOldProc = "SearchTransaction"
   'On Error Goto errProc
   SearchTransaction = False
   
   If pbInitTran = False Then GoTo endProc
   
   If Not IsMissing(sSearch) Then
      lsValue = sSearch
      lsField = "a.sTransNox"
      lsDescript = ""
      If Not IsMissing(bByCode) Then
         If bByCode = False Then
            lsField = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
            lsDescript = "CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)"
         End If
      End If
   End If
   
   SearchTransaction = p_oBaseClas.SearchTransaction(lsValue, lsField, lsDescript)
   
endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & sSearch _
                        & ", " & bByCode & " )"
End Function

Function LoadDetail() As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "LoadDetail"
   'On Error Goto errProc
   LoadDetail = False
   
   With p_oBaseClas
      If .Master("sTransNox") = Empty Then GoTo endProc
      
      Set lors = New Recordset
   
      lsSQL = "SELECT" & _
                  "  c.sAcsModID" & _
                  ", SUM(c.nSQtyIsue) xQuantity" & _
                  ", c.nQtyOnHnd" & _
               " FROM MC_SO_Detail a" & _
                  ", MC_Serial b" & _
                  ", Access_Model c" & _
               " WHERE a.sSerialID = b.sSerialID" & _
                  " AND b.sModelIDx = c.sModelIDx" & _
                  " AND a.sTransNox = " & strParm(.Master("sTransNox")) & _
                  " AND a.cMotorNew = " & strParm(1) & _
                  " AND c.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND c.sRecdStat = " & strParm(xeRecStateActive) & _
               " GROUP BY c.sAcsModID" & _
                  ", c.nQtyOnHnd"
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If lors.EOF Then GoTo endProc
      
      pnCtr = 0
      Do
         If pnCtr > .ItemCount - 1 Then .AddDetail
         
         .Detail(pnCtr, "sTransNox") = .Master("sTransNox")
         .Detail(pnCtr, "sDescript") = lors("sDescript")
         .Detail(pnCtr, "nQtyOnHnd") = lors("nQtyOnHnd")
         .Detail(pnCtr, "nQuantity") = lors("xQuantity")
         .Detail(pnCtr, "nGivenxxx") = 0
         .Detail(pnCtr, "sAcsModID") = lors("sAcsModID")
         
         pnCtr = pnCtr + 1
         lors.MoveNext
      Loop Until lors.EOF
   End With
   LoadDetail = True
   
endProc:
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Function UpdateTransaction() As Boolean
   If pbInitTran = False Then Exit Function
   
   p_oBaseClas.EditMode = xeModeUpdate
   
   UpdateTransaction = True
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, "clsAccSales", lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub

Private Sub getDetail(ByVal lnRow As Integer, _
                      ByVal lsValue As String, _
                      ByVal lbSearch As Boolean)
   Dim lsDetail As String
   Dim lasDetail() As String
   Dim lsOldProc As String

   lsOldProc = "getDetail"
   'On Error Goto errProc
   
   With p_oBaseClas
      If lsValue <> "" Then
         If lsValue = .Detail(lnRow, "sDescript") Then GoTo endProc
         If lbSearch Then
            lsDetail = "sDescript LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsDetail = "sDescript = " & strParm(Trim(lsValue))
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
      .Detail(lnRow, "sAcsModID") = lasDetail(0)
      .Detail(lnRow, "sDescript") = lasDetail(1)
      .Detail(lnRow, "nQtyOnHnd") = lasDetail(2)
   End With
   
   RaiseEvent DetailRetrieved(2)

endProc:
   Exit Sub
endWithClear:
   With p_oBaseClas
      .Detail(lnRow, "sAcsModID") = ""
      .Detail(lnRow, "sDescript") = ""
      .Detail(lnRow, "nQtyOnHnd") = 0
   End With
   RaiseEvent DetailRetrieved(2)
   GoTo endProc
errProc:
    ShowError lsOldProc & "( " & lnRow _
                        & ", " & lsValue _
                        & ", " & lbSearch & " )"
End Sub

Private Function UndoTransaction() As Boolean
   Dim loAccessory As clsAccessory
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "UndoTransaction"
   'On Error Goto errProc
   UndoTransaction = False
   
   Set loAccessory = New clsAccessory
   Set loAccessory.AppDriver = p_oAppDrivr
   
   With p_oBaseClas
      Set lors = New Recordset
      lsSQL = AddCondition(.DetailQuery, _
                  "a.sTransNox = " & strParm(.Master("sTransNox")))
      
      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
      If lors.EOF Then
         UndoTransaction = True
         GoTo endProc
      End If
      
      Do
         If lors("nGivenxxx") > 0 Then
            loAccessory.InitTransaction
            loAccessory.TransactDate = .Master("dTransact")
            loAccessory.Branch = p_sBranchCd
            loAccessory.Quantity = lors("nGivenxxx")
            loAccessory.AccessID = lors("sAcsModID")
            If loAccessory.Sales(.Master("sTransNox"), True) = False Then GoTo endProc
         End If
         lors.MoveNext
      Loop Until lors.EOF
   End With
   UndoTransaction = True
   
endProc:
   Set loAccessory = Nothing
   Set lors = Nothing
   Exit Function
errProc:
    ShowError lsOldProc & "( " & " )"
End Function

Private Sub p_oBaseClas_LoadOthers()
   With p_oBaseClas
      For pnCtr = 0 To .ItemCount
         .Detail(pnCtr, "nQtyOnHnd") = .Detail(pnCtr, "xQtyOnHnd")
      Next
   End With
End Sub

Private Sub p_oBaseClas_Save(Saved As Boolean)
   Saved = True
End Sub

Private Sub p_oBaseClas_WillSave(Cancel As Boolean)
   Dim loAccessory As clsAccessory
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lnCtr As Long
   
   lsOldProc = "p_oBaseClas_WillSave"
   'On Error Goto errProc
   Cancel = True
   
   pbModified = True
   With p_oBaseClas
      If .isDetailModify = False Then
         pbModified = False
         Cancel = False
         GoTo endProc
      End If
      
      If .EditMode = xeModeUpdate Then
         If UndoTransaction = False Then GoTo endProc
         
         lsSQL = "DELETE FROM MC_SO_Accessories" & _
                  " WHERE sTransNox = " & strParm(.Master("sTransNox"))
                           
         p_oAppDrivr.RegisDelete lsSQL
         p_oAppDrivr.Connection.Execute lsSQL, lnRow, adCmdText
      End If
         
      Set loAccessory = New clsAccessory
      Set loAccessory.AppDriver = p_oAppDrivr
      loAccessory.InitTransaction
      loAccessory.TransactDate = .Master("dTransact")
      loAccessory.Branch = p_sBranchCd
      
      lnCtr = 1
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "nQuantity") > 0 Then
            If .Detail(pnCtr, "nGivenxxx") > 0 Then
               loAccessory.AccessID = .Detail(pnCtr, "sAcsModID")
               loAccessory.Quantity = .Detail(pnCtr, "nGivenxxx")
               
               If loAccessory.Sales(.Master("sTransNox")) = False Then GoTo endProc
            End If
            
            lsSQL = "INSERT INTO MC_SO_Accessories " & _
                        "( sTransNox" & _
                        ", nEntryNox" & _
                        ", sAcsModID" & _
                        ", nQuantity" & _
                        ", nGivenxxx" & _
                        ", dModified" & _
                     " ) VALUES ( " & _
                        strParm(.Master("sTransNox")) & _
                        ", " & lnCtr & _
                        ", " & strParm(.Detail(pnCtr, "sAcsModID")) & _
                        ", " & .Detail(pnCtr, "nQuantity") & _
                        ", " & .Detail(pnCtr, "nGivenxxx") & _
                        ", " & dateParm(p_oAppDrivr.ServerDate()) & ")"
            p_oAppDrivr.Connection.Execute lsSQL, lnRow, adCmdText
            
            If lnRow = 0 Then
               MsgBox "Unable to Update MC Sales Accessories!!!", vbCritical, "Warning"
               GoTo endProc
            End If
            
            lnCtr = lnCtr + 1
         End If
      Next
   End With
   
   Cancel = False
   
endProc:
   Set loAccessory = Nothing
   Exit Sub
errProc:
    ShowError lsOldProc & "( " & Cancel & " )"
End Sub
