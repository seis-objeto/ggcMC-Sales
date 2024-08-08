VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccidentInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Accident Insurance for Purchase of Motorcycle
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
'  iMac[ 10/29/2015 02:05 pm ]
'     Start creating this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "clsAccidentInsurance"
Private Const pxeMasTable As String = "MC_Accident_Insurance"
Private Const pxeDetTable As String = "MC_Accident_Insurance_Beneficiaries"
Private Const pxePolicyTp As String = "Accident_Insurance"
Private Const pxeDEFAULTINSURANCE = "0001"

Private p_oAppDrivr As clsAppDriver

Private p_sBranchCd As String
Private p_sParentxx As String
Private p_sSourceCd As String

Private p_oMaster As Recordset
Private p_oDetail As Recordset

Private p_oClient As clsNeoClient 'insured object
Private p_oAssClt As clsNeoClient
Private p_oBenft1 As clsNeoClient '1st beneficiary object
Private p_oBenft2 As clsNeoClient '2nd beneficiary object
Private p_oBenft3 As clsNeoClient '3rd beneficiary object

Private p_sMasterSQL As String
Private p_sDetailSQL As String
Private p_sBrowseSQL As String

Private psClientNm As String 'insured name

Private p_sBeneft1 As String '1st beneficiary name
Private p_sBeneft2 As String '2nd beneficiary name
Private p_sBeneft3 As String '3rd beneficiary name
Private p_sOrgBnft As String 'stores the original clientid of beneficiaries
Private p_sORNoxxxx As String

Private p_nEditMode As xeEditMode

Private pbInitTran As Boolean
Private pbHsMaster As Boolean
Private pbShowMsg As Boolean
Private pbTeleMktg As Boolean

Private pnCtr As Integer
Private pnRow As Integer

Public Event MasterRetreive(ByVal Index As Integer, ByVal Value As Variant)
Public Event DetailRetreive(ByVal Index As Integer, ByVal Value As Variant)
Public Event OthersRetreive(ByVal Index As Integer, ByVal Value As Variant)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set Client(Value As clsNeoClient)
   Set p_oClient = Value
   Set p_oAssClt = xCopy(p_oClient)
   p_oMaster("sClientID") = p_oAssClt.Master("sClientID")
   psClientNm = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
   
   p_oMaster("xFullName") = p_oClient.Master("sLastName") + ", " + p_oClient.Master("sFrstName") + " " + Trim(p_oClient.Master("sSuffixNm")) + IIf(Trim(p_oClient.Master("sSuffixNm")) = "", "", " ") + p_oClient.Master("sMiddName")
   p_oMaster("xAddressx") = IIf(Trim(p_oClient.Master("sHouseNox")) = "", "", p_oClient.Master("sHouseNox") & " ") & p_oClient.Master("sAddressx") & ", " & p_oClient.Master("sTownName")
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Benificiary1(ByVal Value As String)
   p_sBeneft1 = Value
End Property

Property Let Benificiary2(ByVal Value As String)
   p_sBeneft2 = Value
End Property

Property Let Benificiary3(ByVal Value As String)
   p_sBeneft3 = Value
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Let ORNo(ByVal Value As String)
   p_sORNoxxxx = Value
End Property

Property Let DisplayConfirmation(Value As Boolean)
   pbShowMsg = Value
End Property

Property Let showMessage(Value As Boolean)
   pbShowMsg = Value
End Property

Property Get ItemCount()
   ItemCount = p_oDetail.RecordCount
End Property

Property Get EditMode()
   EditMode = p_nEditMode
End Property

Property Let Master(ByVal Index As Variant, ByVal Value As Variant)
   If Not IsNumeric(Index) Then Index = LCase(Index)
   
   Select Case Index
   Case 3, "sclientid"
      p_oMaster(Index) = Value
   Case 8, "ssourceno"
      If Len(Value) > 12 Then
         If Mid(Value, 13) = "TLMi" Or Mid(Value, 13) = "TLMc" Then
            p_oMaster(Index) = Mid(Value, 1, 12)
            
            p_sSourceCd = Mid(Value, 13)
            pbTeleMktg = True
         Else
            p_oMaster(Index) = Value
         End If
      Else
         p_oMaster(Index) = Value
      End If
   Case Else
      p_oMaster(Index) = Value
   End Select
End Property

Property Get Master(ByVal Index As Variant)
   Master = p_oMaster(Index)
End Property

Property Let Detail(ByVal Row As Integer, ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0 'i didn't follow the standards on this index
      getBenefit1 Value, True
   Case 2 'i didn't follow the standards on this index
      getBenefit2 Value, True
   Case 4 'i didn't follow the standards on this index
      getBenefit3 Value, True
   Case 5, "sRelatnID"
      p_oDetail.Move Row, adBookmarkFirst
      getRelation Row, Value, True
   Case "sClientID"
      p_oDetail.Move Row, adBookmarkFirst
      p_oDetail("sClientID") = Value
   End Select
End Property

Property Get Detail(ByVal Row As Integer, ByVal Index As Variant)
   p_oDetail.Move Row, adBookmarkFirst
   
   Detail = p_oDetail(Index)
End Property

Function InitTransaction() As Boolean
   Dim lsOldProc As String
   
   lsOldProc = pxeMODULENAME & ".InitTransaction"
   Debug.Print lsOldProc
   
   'On Error GoTo errProc
   
   If isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Application Driver is Not Initialized for this Object!" & vbCrLf & _
            "Set Application Driver First then Try Again!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   p_sMasterSQL = "SELECT" & _
                     "  a.sTransNox" & _
                     ", a.dTransact" & _
                     ", a.dSalesxxx" & _
                     ", a.sClientID" & _
                     ", a.sSerialID" & _
                     ", a.sInsPrmID" & _
                     ", a.sGCardNox" & _
                     ", a.sOthrInfo" & _
                     ", a.sSourceNo" & _
                     ", a.sSourceCd" & _
                     ", a.cTranStat" & _
                     ", a.sModified" & _
                     ", a.dModified" & _
                     ", b.sCompnyNm xFullName" & _
                     ", a.sOthrInfo xAddressx" & _
                     ", c.sDescript xInsurDsc" & _
                     ", c.nUpSalesx" & _
                  " FROM " & pxeMasTable & " a" & _
                     " LEFT JOIN Client_Master b" & _
                        " ON a.sClientID = b.sClientID" & _
                     " LEFT JOIN Accident_Insurance c" & _
                        " ON a.sInsPrmID = c.sInsPrmID"
                  
                  
   Set p_oMaster = New Recordset
   p_oMaster.Open AddCondition(p_sMasterSQL, "0=1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oMaster.ActiveConnection = Nothing
                  
   p_sDetailSQL = "SELECT" & _
                     "  a.sClientID" & _
                     ", a.sTransNox" & _
                     ", a.dModified" & _
                     ", b.sCompnyNm xClientNm" & _
                     ", b.dBirthDte" & _
                     ", a.sRelatnID" & _
                     ", c.sRelatnDs" & _
                  " FROM " & pxeDetTable & " a" & _
                     " LEFT JOIN Client_Master b" & _
                        " ON a.sClientID = b.sClientID" & _
                     " LEFT JOIN Relation c" & _
                        " ON a.sRelatnID = c.sRelatnID"
                     
   Set p_oDetail = New Recordset
   p_oDetail.Open AddCondition(p_sDetailSQL, "0=1"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oDetail.ActiveConnection = Nothing
                  
   p_sBrowseSQL = "SELECT" & _
                     "  sInsPrmID" & _
                     ", sDescript" & _
                     ", cRecdStat" & _
                  " FROM " & pxePolicyTp
                  
   Set p_oClient = New clsNeoClient
   With p_oClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oAssClt = New clsNeoClient
   With p_oAssClt
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oBenft1 = New clsNeoClient
   With p_oBenft1
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oBenft2 = New clsNeoClient
   With p_oBenft2
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With
   
   Set p_oBenft3 = New clsNeoClient
   With p_oBenft3
      Set .AppDriver = p_oAppDrivr
      .Branch = p_sBranchCd
      If .InitClient() = False Then GoTo endProc
   End With

   pbInitTran = True
   InitTransaction = True
   p_nEditMode = xeModeReady
endProc:
   Exit Function
errProc:
   Call ShowError(lsOldProc)
End Function

Function NewTransaction() As Boolean
   If Not pbInitTran Then GoTo endProc
   
   pnRow = 0
   
   Call initMaster
   Call InitDetail
   
   pbHsMaster = False
   NewTransaction = True
   p_nEditMode = xeModeAddNew
endProc:
   Exit Function
End Function

Function UpdateTransaction() As Boolean
   If Not pbInitTran Then GoTo endProc
   
   'dont allow updates on posted or cancelled transactions
   If p_oMaster("cTranStat") <> xeStateOpen And p_oMaster("cTranStat") <> xeStateClosed Then
      MsgBox "Unable to Update Posted/Cancelled Records.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_nEditMode = xeModeUpdate
   UpdateTransaction = True
endProc:
   Exit Function
End Function

Private Function saveClients() As Boolean
   If p_oBenft1.Master("sLastName") <> "" Then
      If p_oBenft1.SaveClient = False Then
         MsgBox "Benefeciary 1 info not saved and will be remove!" & vbCrLf & _
                         "Please make sure to re-enter co-client info.", vbCritical, "Warning"
         GoTo endProc
      Else
         Detail(0, "sClientID") = p_oBenft1.Master("sClientID")
      End If
   End If
   
   If p_oBenft2.Master("sLastName") <> "" Then
      If p_oBenft2.SaveClient = False Then
         MsgBox "Benefeciary 2 info not saved and will be remove!" & vbCrLf & _
                         "Please make sure to re-enter co-client info.", vbCritical, "Warning"
         GoTo endProc
      Else
         Detail(1, "sClientID") = p_oBenft2.Master("sClientID")
      End If
   End If
   
   If p_oBenft3.Master("sLastName") <> "" Then
      If p_oBenft3.SaveClient = False Then
         MsgBox "Benefeciary 3 info not saved and will be remove!" & vbCrLf & _
                         "Please make sure to re-enter co-client info.", vbCritical, "Warning"
         GoTo endProc
      Else
         Detail(2, "sClientID") = p_oBenft3.Master("sClientID")
      End If
   End If
   
   saveClients = True
endProc:
   Exit Function
End Function

Private Function validateDetail(ByRef lsNewDetail As String) As Boolean

   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lsRemove As String
   Dim lsCondition As String
   Dim lasDetail() As String

   If p_sOrgBnft <> "" Then
      For pnCtr = 0 To p_oDetail.RecordCount - 1
         If InStr(1, p_sOrgBnft, Detail(pnCtr, "sClientID"), vbTextCompare) <> 0 Then
            lsDetail = lsDetail & "»" & Detail(pnCtr, "sClientID")
         End If
      Next
      If lsDetail <> "" Then lsDetail = Mid(lsDetail, 2)
         
      lasDetail = Split(lsDetail, "»")
      lsCondition = ""
      For lnCtr = 0 To UBound(lasDetail)
         lsCondition = lsCondition & "," & strParm(lasDetail(lnCtr))
      Next
      
      If lsCondition <> "" Then
         lsCondition = Mid(lsCondition, 2)
         
         lsNewDetail = Replace(lsCondition, ",", "»")
         lsCondition = "(" & lsCondition & ")"
      Else
         lsCondition = ""
      End If
      
      lsSQL = "DELETE FROM MC_Accident_Insurance_Beneficiaries" & _
            " WHERE sTransNox = " & strParm(p_oMaster("sTransNox")) & _
               IIf(lsCondition <> "", " AND sClientID NOT IN " & lsCondition, "")
            
      p_oAppDrivr.Execute lsSQL, "MC_Accident_Insurance_Beneficiaries"
   End If

   validateDetail = True
End Function

Function SaveTransaction() As Boolean
   Dim lsSQL As String
   Dim lsDetail As String
   Dim lsOldProc As String
   Dim lbIsModified As Boolean
   
   Dim lbApproved As Boolean
   Dim lnUserRights As Integer
   Dim lsUserID As String, lsUserName As String
   
   lsOldProc = pxeMODULENAME & ".SaveTransaction"
   Debug.Print lsOldProc
   
   'On Error GoTo errProc
   
   If pbShowMsg Then p_oAppDrivr.BeginTrans
   
   If Not pbInitTran Then GoTo endProc
      
   'kalyptus - 2016.06.18 03:02pm
   'If there is a difference between the assured individual and the mc client then
   'ignore the client id of the mc client.
   'Note: mc client was set as the default assured client and his/her client id
   '      is re-assigned before calling this method/function/procedure.
   If Not (p_oAssClt.Master("sLastName") = p_oClient.Master("sLastName") And _
           p_oAssClt.Master("sFrstName") = p_oClient.Master("sFrstName") And _
           p_oAssClt.Master("sMiddName") = p_oClient.Master("sMiddName") And _
           p_oAssClt.Master("sSuffixNm") = p_oClient.Master("sSuffixNm")) Then
       
         If Not p_oAssClt.SaveClient Then
            MsgBox "Unable to Save Assured client.", vbCritical, "Warning"
            GoTo endProc
         End If
         p_oMaster("sClientID") = p_oAssClt.Master("sClientID")
   End If
   
   If Not isEntryOK Then GoTo endProc
   If Not saveClients Then GoTo endProc
   
   If p_nEditMode = xeModeAddNew Or Not pbHsMaster Then
      If pbTeleMktg = False Then
         If p_oMaster("nUpSalesx") > 0 Then If Not saveReceipt Then GoTo endProc
      Else
         'iMac 2018-12-11
         'singilin ang customer pag upsale pag installment ang payment form
         If p_sSourceCd = "TLMi" Then
            If p_oMaster("nUpSalesx") > 0 Then If Not saveReceipt Then GoTo endProc
         End If
         
         p_oMaster("sOthrInfo") = p_oMaster("sOthrInfo") + "(TLMC)"
      End If
   ElseIf p_nEditMode = xeModeUpdate Then
      If Not validateDetail(lsDetail) Then GoTo endProc
   End If
      
   Select Case p_nEditMode
   Case xeModeAddNew, xeModeUpdate
      If Not pbHsMaster Then
         lsSQL = ADO2SQL(p_oMaster, "MC_Accident_Insurance", , _
                           p_oAppDrivr.UserID, p_oAppDrivr.ServerDate, "xFullName»xInsurDsc»nUpSalesx»xAddressx")

         If p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance") = 0 Then
            MsgBox "Unable to Save Assured Info.", vbCritical, "Warning"
            GoTo endProc
         End If
         lbIsModified = True
      End If
      
      p_oDetail.MoveFirst
      Do Until p_oDetail.EOF
         If Trim(p_oDetail("sClientID")) <> "" Then
            Select Case p_nEditMode
            Case xeModeAddNew
               lsSQL = ADO2SQL(p_oDetail, "MC_Accident_Insurance_Beneficiaries", , , _
                              p_oAppDrivr.ServerDate, "xClientNm»sRelatnDs»xAgexxxxx»dBirthDte")
         
               If p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance_Beneficiaries") = 0 Then
                  MsgBox "Unable to Save Assured Detail.", vbCritical, "Warning"
                  GoTo endProc
               End If
               lbIsModified = True
            Case xeModeUpdate
               If InStr(1, lsDetail, p_oDetail("sClientID"), vbTextCompare) = 0 Then
                  lsSQL = ADO2SQL(p_oDetail, "MC_Accident_Insurance_Beneficiaries", , , _
                                 p_oAppDrivr.ServerDate, "xClientNm»sRelatnDs»xAgexxxxx»dBirthDte")
            
                  If p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance_Beneficiaries") = 0 Then
                     MsgBox "Unable to Save Assured Detail.", vbCritical, "Warning"
                     GoTo endProc
                  End If
                  lbIsModified = True
               End If
            End Select
         End If
         p_oDetail.MoveNext
      Loop
   End Select
   
   If pbShowMsg And lbIsModified = True Then
      lbApproved = True
      If p_oAppDrivr.UserLevel < xeManager Then
         MsgBox "Updating of Insurance Info requires Managers Approval", vbInformation + vbOKOnly
         If Not GetApproval(p_oAppDrivr, lnUserRights, lsUserID, lsUserName, p_oAppDrivr.MenuName) Then GoTo endProc
      Else
         lnUserRights = p_oAppDrivr.UserLevel
      End If
   
      If lnUserRights < xeManager Then
         MsgBox "Updating of Insurance Info requires MANAGER account..." & vbCrLf & _
                "Can't proceed because of insufficient RIGHT!", vbInformation + vbOKOnly
         lbApproved = False
      End If
      
      If lbApproved Then
         MsgBox "Record Saved Successfuly.", vbInformation, "Success"
      Else
         MsgBox "Unable to Update Record.", vbCritical, "Warning"
         GoTo endWithRoll
      End If
   End If
   If pbShowMsg Then p_oAppDrivr.CommitTrans
      
   SaveTransaction = True
endProc:
   Exit Function
endWithRoll:
   If pbShowMsg Then p_oAppDrivr.RollbackTrans
   Exit Function
errProc:
   Call ShowError(lsOldProc)
End Function

Function SearchMaster(ByVal Index As Variant, ByVal Value As Variant, ByVal lbSearch As Boolean, ByVal lbByCode As Boolean)
   Select Case Index
   Case 5, "sInsPrmID"
      Call getInsurance(Value, lbSearch, lbByCode)
   End Select
End Function

Private Function isValidBDay() As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   
   'iMac 2016-07-28
   '  validate only date on adding new transaction,
   '  since update of the insured person was not applicable on registers.
   '     birthdate of insured person becomes 1900-01-01, when we are updating the benefeciary.
   If p_nEditMode = xeModeAddNew Then
      If DateDiff("d", p_oAssClt.Master("dBirthDte"), p_oAppDrivr.SysDate) / 365 < 18 Or _
         DateDiff("d", p_oAssClt.Master("dBirthDte"), p_oAppDrivr.SysDate) / 365 >= 61 Then
         
         GoTo endProc
      End If
   End If
   
   isValidBDay = True
endProc:
   Exit Function
End Function

Private Function isEntryOK() As Boolean
   If p_oMaster("sClientID") = "" Then
      MsgBox "Invalid Client Detected.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Not isValidBDay Then
      MsgBox "Invalid Age for Insured Person. Only 18 to 60 years old is accepted!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_oMaster("sSerialID") = "" Then
      MsgBox "Invalid Motorcycle Serial Detected.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_oMaster("sInsPrmID") = "" Then
      MsgBox "Invalid Insurance Coverage Detected.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_oMaster("sSourceNo") = "" And p_oMaster("sSourceCd") = "" Then
      MsgBox "No Source Detected.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   p_oDetail.MoveFirst
   Do Until p_oDetail.EOF
      If Trim(p_oDetail("sClientID")) <> "" And Trim(p_oDetail("sRelatnID")) = "" Then
         MsgBox "Invalid Relation ID Detected.", vbCritical, "Warning"
         GoTo endProc
      End If
   
      p_oDetail.MoveNext
   Loop
   
   isEntryOK = True
endProc:
   Exit Function
End Function

Private Function isDetailModify() As Boolean
   With p_oDetail
      .MoveFirst
      If .RecordCount = 1 And Trim(.Fields("sClientID")) = "" And pnRow = 0 Then
         isDetailModify = False
         GoTo endProc
      End If
      
      If .RecordCount = pnRow Then
         Do Until .EOF
            If p_oDetail("sClientID").OriginalValue <> p_oDetail("sClientID") Then
            End If
            If Trim(p_oDetail("sClientID")) = "" Then
               isDetailModify = True
               GoTo endProc
            End If
            
            If .Fields("sRelatnID").OriginalValue <> .Fields("sRelatnID") Then
               isDetailModify = True
               GoTo endProc
            End If
            .MoveNext
         Loop
      ElseIf .RecordCount <> pnRow Then
         isDetailModify = True
         GoTo endProc
      End If
      
      isDetailModify = False
   End With
endProc:
   Exit Function
End Function

Function getClient(ByVal lsValue As String) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getClient"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'On Error GoTo errProc
   
   'Load client record
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
   End With
   
   'kalyptus - 2016.06.18 03:07pm
   'Instead of leaving the client info as blank if user set the name to blank
   'reload the value of p_oAssClt instead...
   If lsValue <> "" Then
      If Trim(LCase(lsValue)) = Trim(LCase(psClientNm)) Then GoTo endProc
      lbExist = loClient.SearchClient(lsValue, False)

      If Not lbExist Then
         lasName = GetSplitedName(lsValue)
         loClient.Master("sLastName") = lasName(0)
         loClient.Master("sFrstName") = lasName(1)
      End If
   
      If loClient.getClient Then
         Set p_oAssClt = loClient
         p_oAssClt.Master("cMCClient") = "1"
      Else
         Set p_oAssClt = xCopy(p_oClient)
      End If
   End If
      
   p_oMaster("sClientID") = p_oAssClt.Master("sClientID")
   p_oMaster("xFullName") = p_oAssClt.Master("sLastName") + ", " + p_oAssClt.Master("sFrstName") + " " + Trim(p_oAssClt.Master("sSuffixNm")) + IIf(Trim(p_oAssClt.Master("sSuffixNm")) = "", "", " ") + p_oAssClt.Master("sMiddName")
   p_oMaster("xAddressx") = IIf(Trim(p_oAssClt.Master("sHouseNox")) = "", "", p_oAssClt.Master("sHouseNox") & " ") & p_oAssClt.Master("sAddressx") & ", " & p_oAssClt.Master("sTownName")
   
   RaiseEvent OthersRetreive(0, p_oMaster("sClientID"))
   RaiseEvent OthersRetreive(1, p_oMaster("xFullName"))
   RaiseEvent OthersRetreive(2, p_oMaster("xAddressx"))
   RaiseEvent OthersRetreive(3, Format(p_oAssClt.Master("dBirthDte"), "Mmm dd, yyyy"))
   RaiseEvent OthersRetreive(4, DateDiff("yyyy", p_oAssClt.Master("dBirthDte"), p_oAppDrivr.SysDate))
   RaiseEvent OthersRetreive(6, IFNull(p_oAssClt.Master("sOccptnNm"), "N-O-N-E"))
   RaiseEvent OthersRetreive(7, p_oAssClt.Master("sMobileNo"))
   
   Select Case p_oAssClt.Master("cCvilStat")
   Case "0"
      RaiseEvent OthersRetreive(5, "SINGLE")
   Case "1"
      RaiseEvent OthersRetreive(5, "MARRIED")
   Case "2"
      RaiseEvent OthersRetreive(5, "SEPARATED")
   Case "3"
      RaiseEvent OthersRetreive(5, "WIDOWED")
   End Select
   
   getClient = True
   
endProc:
   Exit Function
errProc:
    ShowError lsProcName & "( " & lsValue & " )"
End Function

Function DeleteDetail(ByVal lnRow As Integer) As Boolean
   If Not pbInitTran Then Exit Function
   
   p_oDetail.Move lnRow, adBookmarkFirst
   p_oDetail.Delete adAffectCurrent
   
   DeleteDetail = True
End Function

Private Function getBenefit1(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getBenefit1"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'On Error GoTo errProc
   
   p_oDetail.Move 0, adBookmarkFirst
   
   If lsValue = "" Then GoTo endWithClear
   If p_oBenft1.Master("sCompnyNm") = lsValue Or p_sBeneft1 = lsValue Then
      getBenefit1 = True
      GoTo endProc
   End If
   
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
   End With
   
   If lsValue <> "" Then
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
      Set p_oBenft1 = loClient
   End If
   
   p_oBenft1.Master("cMCClient") = "1"
   p_oDetail("sClientID") = loClient.Master("sClientID")
   
   p_oDetail("sRelatnID") = ""
   RaiseEvent DetailRetreive(80, "")

   getBenefit1 = True
endProc:
   RaiseEvent DetailRetreive(0, IIf(p_oBenft1.Master("sCompnyNm") = "", p_sBeneft1, p_oBenft1.Master("sCompnyNm")))
   Exit Function
endWithClear:
   p_oBenft1.InitClient
   p_oDetail("sClientID") = ""
   p_oDetail("xClientNm") = ""
   p_oDetail("sRelatnID") = ""
   
   RaiseEvent DetailRetreive(0, "")
   RaiseEvent DetailRetreive(80, "")
   Exit Function
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getBenefit2(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getBenefit2"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'On Error GoTo errProc
   
   p_oDetail.Move 1, adBookmarkFirst
   
   If lsValue = "" Then GoTo endWithClear
   If p_oBenft2.Master("sCompnyNm") = lsValue Or p_sBeneft2 = lsValue Then
      getBenefit2 = True
      GoTo endProc
   End If
   
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
   End With
   
   If lsValue <> "" Then
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
      Set p_oBenft2 = loClient
   End If
   
   p_oBenft2.Master("cMCClient") = "1"
   p_oDetail("sClientID") = loClient.Master("sClientID")
   
   p_oDetail("sRelatnID") = ""
   RaiseEvent DetailRetreive(81, "")

   getBenefit2 = True
endProc:
   RaiseEvent DetailRetreive(2, IIf(p_oBenft2.Master("sCompnyNm") = "", p_sBeneft2, p_oBenft2.Master("sCompnyNm")))
   Exit Function
endWithClear:
   p_oBenft2.InitClient
   p_oDetail("sClientID") = ""
   p_oDetail("xClientNm") = ""
   p_oDetail("sRelatnID") = ""
   
   RaiseEvent DetailRetreive(2, "")
   RaiseEvent DetailRetreive(81, "")
   Exit Function
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getBenefit3(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lsProcName As String
   Dim lasName() As String
   Dim lbExist As Boolean
   Dim loClient As clsNeoClient

   lsProcName = "getBenefit3"
   Debug.Print pxeMODULENAME & "." & lsProcName
   
   'On Error GoTo errProc
   p_oDetail.Move 2, adBookmarkFirst
   
   If lsValue = "" Then GoTo endWithClear
   If p_oBenft3.Master("sCompnyNm") = lsValue Or p_sBeneft3 = lsValue Then
      getBenefit3 = True
      GoTo endProc
   End If
   
   Set loClient = New clsNeoClient
   With loClient
      Set .AppDriver = p_oAppDrivr
      .Branch = p_oAppDrivr.BranchCode
      If .InitClient() = False Then GoTo endProc
   End With
   
   If lsValue <> "" Then
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
      Set p_oBenft3 = loClient
   End If
   
   p_oBenft3.Master("cMCClient") = "1"
   p_oDetail("sClientID") = loClient.Master("sClientID")
   
   p_oDetail("sRelatnID") = ""
   RaiseEvent DetailRetreive(82, "")

   getBenefit3 = True
endProc:
   RaiseEvent DetailRetreive(4, IIf(p_oBenft3.Master("sCompnyNm") = "", p_sBeneft3, p_oBenft3.Master("sCompnyNm")))
   Exit Function
endWithClear:
   p_oBenft3.InitClient
   p_oDetail("sClientID") = ""
   p_oDetail("xClientNm") = ""
   p_oDetail("sRelatnID") = ""
   
   RaiseEvent DetailRetreive(4, "")
   RaiseEvent DetailRetreive(82, "")
   Exit Function
errProc:
    ShowError lsProcName & "( " & lsValue _
                        & ", " & lbSearch & " )"
End Function

Private Function getRelation(ByVal Row As Integer, ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsDesript As String
   Dim lsProcName As String
   Dim lasMaster() As String

   lsProcName = "getRelation"
   'On Error GoTo errProc
   
   Select Case Row
   Case 0
      If p_sBeneft1 = lsValue Then
         lsDesript = p_sBeneft1
         GoTo endProc
      End If
   Case 1
      If p_sBeneft2 = lsValue Then
         lsDesript = p_sBeneft2
         GoTo endProc
      End If
   Case 2
      If p_sBeneft3 = lsValue Then
         lsDesript = p_sBeneft3
         GoTo endProc
      End If
   End Select
   
   If lsValue <> "" Then
      If lbSearch Then
         lsSQL = "sRelatnDs LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsSQL = "sRelatnDs = " & strParm(Trim(lsValue))
      End If
   Else
      GoTo endWithClear
   End If
   
   lsSQL = "SELECT sRelatnID" & _
               ", sRelatnDs" & _
            " FROM Relation" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
               IIf(lsSQL = "", "", " AND " & lsSQL) & _
            " ORDER BY sRelatnDs"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   With loRS
      If .RecordCount = 0 Then
         p_oDetail("sRelatnID") = ""
         p_oDetail("sRelatnDs") = ""
      ElseIf .RecordCount = 1 Then
         p_oDetail("sRelatnID") = loRS("sRelatnID")
         p_oDetail("sRelatnDs") = loRS("sRelatnDs")
      Else
         lsSQL = KwikBrowse(p_oAppDrivr, loRS _
                              , "sRelatnID»sRelatnDs" _
                              , "ID»Relation")
         If lsSQL = "" Then
            p_oDetail("sRelatnID") = ""
            p_oDetail("sRelatnDs") = ""
            GoTo endProc
         Else
            lasMaster = Split(lsSQL, "»")
            p_oDetail("sRelatnID") = lasMaster(0)
            p_oDetail("sRelatnDs") = lasMaster(1)
         End If
      End If
   End With
   getRelation = True
endProc:
   Select Case Row
   Case 0
      p_sBeneft1 = lsDesript
   Case 1
      p_sBeneft2 = lsDesript
   Case 2
      p_sBeneft3 = lsDesript
   End Select


   RaiseEvent DetailRetreive(Row + 80, p_oDetail("sRelatnDs"))

   Set loRS = Nothing
   Exit Function
endWithClear:
   p_oDetail("sRelatnID") = ""
   p_oDetail("sRelatnDs") = ""
   
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lbSearch & " )"
End Function

Private Function getInsurance(ByVal lsValue As String, _
                                 ByVal lbSearch As Boolean, _
                                 Optional ByVal lbByCode As Boolean = False) As Boolean
                                 
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lasMaster() As String
   Dim lsProcName As String

   lsProcName = "getRelation"
   'On Error GoTo errProc
   
   If lsValue <> "" Then
      If Not lbByCode Then
         If lbSearch Then
            lsSQL = "sDescript LIKE " & strParm(Trim(lsValue) & "%")
         Else
            lsSQL = "sDescript = " & strParm(Trim(lsValue))
         End If
      Else
         lsSQL = "sInsPrmID = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endProc
   End If
   
   lsSQL = "SELECT sInsPrmID" & _
               ", sDescript" & _
               ", nUpSalesx" & _
            " FROM Accident_Insurance" & _
            " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
               IIf(lsSQL = "", "", " AND " & lsSQL) & _
            " ORDER BY sDescript"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   If loRS.EOF Then
      RaiseEvent MasterRetreive(5, p_oMaster("xInsurDsc"))
      GoTo endProc
   End If
   
   If loRS.RecordCount > 1 Then
      lsSQL = KwikBrowse(p_oAppDrivr, loRS _
                           , "sInsPrmID»sDescript" _
                           , "ID»Coverage")
      
      If lsSQL <> "" Then
         lasMaster = Split(lsSQL, "»")
         p_oMaster("sInsPrmID") = lasMaster(0)
         p_oMaster("xInsurDsc") = lasMaster(1)
         p_oMaster("nUpSalesx") = lasMaster(2)
         
         RaiseEvent MasterRetreive(5, lasMaster(1))
         GoTo endProc
      End If
   End If
      
   p_oMaster("sInsPrmID") = loRS("sInsPrmID")
   p_oMaster("xInsurDsc") = loRS("sDescript")
   p_oMaster("nUpSalesx") = loRS("nUpSalesx")
   
   RaiseEvent MasterRetreive(5, loRS("sDescript"))
   
   getInsurance = True
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & lsValue _
                            & ", " & lbSearch _
                            & ", " & lbByCode & " )"
End Function

Sub loadClient()
   Dim loRS As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT" & _
               "  a.sClientID" & _
               ", a.sCompnyNm xClientNm" & _
               ", CONCAT(a.sHouseNox, ' ', a.sAddressx, ', ', b.sTownName, ' ', b.sZippCode, ' ', c.sProvName) xAddressx" & _
               ", a.dBirthDte" & _
               ", ROUND(DATEDIFF(NOW(), a.dBirthDte)/365) xAgexxxxx" & _
               ", a.cCvilStat" & _
               ", a.sMobileNo" & _
               ", d.sOccptnNm" & _
               ", e.sCardNmbr" & _
               ", e.sGCardNox" & _
            " FROM Client_Master a" & _
               " LEFT JOIN TownCity b" & _
                  " ON a.sTownIDxx = b.sTownIDxx" & _
               " LEFT JOIN Province c" & _
                  " ON b.sProvIDxx = c.sProvIDxx" & _
               " LEFT JOIN Occupation d" & _
                  " ON a.sOccptnID = d.sOccptnID" & _
               " LEFT JOIN G_Card_Master e" & _
                  " ON a.sClientID = e.sClientID" & _
            " WHERE a.sClientID = " & strParm(p_oMaster("sClientID"))
   
   Set loRS = New Recordset
   loRS.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If loRS.EOF Then
      psClientNm = ""
      GoTo endWithClear
   End If
   
   psClientNm = loRS("xClientNm")
   p_oMaster("xFullName") = loRS("xClientNm")
   p_oMaster("xAddressx") = loRS("xAddressx")
   
   RaiseEvent OthersRetreive(0, loRS("sClientID"))
   RaiseEvent OthersRetreive(1, loRS("xClientNm"))
   RaiseEvent OthersRetreive(2, loRS("xAddressx"))
   RaiseEvent OthersRetreive(3, Format(loRS("dBirthDte"), "Mmm dd, yyyy"))
   RaiseEvent OthersRetreive(4, loRS("xAgexxxxx"))
   RaiseEvent OthersRetreive(6, IFNull(loRS("sOccptnNm"), "N-O-N-E"))
   RaiseEvent OthersRetreive(7, loRS("sMobileNo"))
'   RaiseEvent OthersRetreive(8, IFNull(loRS("sCardNmbr"), "N-O-N-E"))
   
   Select Case loRS("cCvilStat")
   Case "0"
      RaiseEvent OthersRetreive(5, "SINGLE")
   Case "1"
      RaiseEvent OthersRetreive(5, "MARRIED")
   Case "2"
      RaiseEvent OthersRetreive(5, "SEPARATED")
   Case "3"
      RaiseEvent OthersRetreive(5, "WIDOWED")
   End Select
   
   p_oMaster("sClientID") = loRS("sClientID")
   p_oMaster("sGCardNox") = loRS("sGCardNox")
endProc:
   Exit Sub
endWithClear:
   RaiseEvent OthersRetreive(0, "")
   RaiseEvent OthersRetreive(1, "")
   RaiseEvent OthersRetreive(2, "")
   RaiseEvent OthersRetreive(3, "")
   RaiseEvent OthersRetreive(4, "")
   RaiseEvent OthersRetreive(5, "")
   RaiseEvent OthersRetreive(6, "")
   RaiseEvent OthersRetreive(7, "")
'   RaiseEvent OthersRetreive(8, "")
   
   p_oMaster("sClientID") = ""
   p_oMaster("sGCardNox") = ""
End Sub

Function LoadRecord(ByVal sTransNox As String) As Boolean
   If sTransNox = "" Then GoTo endProc
   
   Set p_oMaster = New Recordset

   p_oMaster.Open AddCondition(p_sMasterSQL, "sSourceNo = " & strParm(sTransNox) & " AND sSourceCd = 'MCSl'"), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oMaster.ActiveConnection = Nothing
   pbHsMaster = Not p_oMaster.EOF
   If Not pbHsMaster Then Call initMaster
   
   Set p_oDetail = New Recordset
   p_oDetail.Open AddCondition(p_sDetailSQL, "sTransNox = " & strParm(p_oMaster("sTransNox"))), p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set p_oDetail.ActiveConnection = Nothing
   
   With p_oDetail
      pnRow = .RecordCount

      pnCtr = 0
      p_sOrgBnft = ""
      If Not .EOF Then
         .MoveFirst
         Do Until .EOF
            Select Case pnCtr
            Case 0
               p_sBeneft1 = IFNull(p_oDetail("xClientNm"))
            Case 1
               p_sBeneft2 = IFNull(p_oDetail("xClientNm"))
            Case 2
               p_sBeneft3 = IFNull(p_oDetail("xClientNm"))
            End Select
            p_sOrgBnft = p_sOrgBnft & "»" & Trim(p_oDetail("sClientID"))
            
            .MoveNext
            pnCtr = pnCtr + 1
         Loop
         If p_sOrgBnft <> "" Then p_sOrgBnft = Mid(p_sOrgBnft, 2)
      End If
   End With
   Call InitDetail
   
   LoadRecord = True
   p_nEditMode = xeModeReady
endProc:
   Exit Function
End Function

Private Sub initMaster()
   p_oMaster.AddNew
   p_oMaster("sTransNox") = GetNextCode("MC_Accident_Insurance", "sTransNox", True, p_oAppDrivr.Connection, True, p_sBranchCd)
   p_oMaster("dTransact") = p_oAppDrivr.ServerDate
   p_oMaster("dSalesxxx") = p_oAppDrivr.ServerDate
   p_oMaster("sClientID") = ""
   p_oMaster("sSerialID") = ""
      
   p_oMaster("sInsPrmID") = pxeDEFAULTINSURANCE
   p_oMaster("xInsurDsc") = "Plan 1-FREE"
   p_oMaster("nUpSalesx") = 0
   
   p_oMaster("sGCardNox") = ""
   p_oMaster("sOthrInfo") = ""
   p_oMaster("sSourceNo") = ""
   p_oMaster("sSourceCd") = ""
   p_oMaster("cTranStat") = xeStateOpen
   
   pbTeleMktg = False
End Sub

Private Function InitDetail()

   For pnCtr = pnRow To 2
      p_oDetail.AddNew
      p_oDetail("sTransNox") = p_oMaster("sTransNox")
      p_oDetail("sClientID") = ""
      p_oDetail("sRelatnID") = ""
      p_oDetail("sRelatnDs") = ""
   Next
   
   p_oDetail.MoveFirst
   
   InitDetail = True
End Function

Private Function saveReceipt() As Boolean
   Dim loOReceipt As clsORReceipt
   Dim loFrm As frmInsUpsale
   Dim lsSQL As String
   
   Set loFrm = New frmInsUpsale
   Set loFrm.AppDriver = p_oAppDrivr
   
   loFrm.ORNo = p_sORNoxxxx
   loFrm.Transact = p_oMaster("dSalesxxx")
   loFrm.ClientName = p_oMaster("xFullName")
   loFrm.Address = p_oMaster("xAddressx")
   loFrm.Amount = p_oMaster("nUpSalesx")
   loFrm.Show 1
      
   If loFrm.IsOkey Then
      Set loOReceipt = New clsORReceipt
      With loOReceipt
         Set .AppDriver = p_oAppDrivr
         .Branch = p_sBranchCd
         .Parent = "clsAccidentInsurance"
               
         If .InitTransaction = False Then GoTo endProc
         If .NewTransaction = False Then GoTo endProc
         
         .Master("dTransact") = p_oMaster("dSalesxxx")
         .Master("sORNoxxxx") = loFrm.ORNo
         .Master("sRemarksx") = "MAFRE"
         .Master("nTranAmtx") = IIf(loFrm.ORNo <> "", p_oMaster("nUpSalesx"), 0)
         .Master("cTranType") = "5"
         .Master("sReferNox") = p_oMaster("sSourceNo")
         .Master("sSourceCd") = "MCSl"
         .Master("sClientID") = p_oMaster("sClientID")
         .Master("sCoCltID1") = ""
         .Master("sCoCltID2") = ""
         .Master("sSystemCD") = "MC"
         .Master("sSerialID") = p_oMaster("sSerialID")
         .Master("sInsTypID") = "0109"
         .Master("cRegisFrm") = 6
         .Checks("sCheckNox") = IIf(loFrm.PRNo <> "", loFrm.CheckNo, "")
         .Checks("sAcctNoxx") = IIf(loFrm.PRNo <> "", loFrm.AccountNo, "")
         .Checks("sBankIDxx") = IIf(loFrm.PRNo <> "", loFrm.BankID, "")
         If loFrm.PRNo <> "" Then
            .Checks("dCheckDte") = loFrm.CheckDate
         End If
         .Checks("nClearing") = IIf(loFrm.PRNo <> "", 0, 0)
         .Checks("nAmountxx") = IIf(loFrm.PRNo <> "", p_oMaster("nUpSalesx"), 0)
         .Checks("bIssuedOR") = loFrm.IsORIssued
         .Master("cTranStat") = 0
         .PRNumber = loFrm.PRNo
         
         If .SaveTransaction = False Then GoTo endProc
      End With
      
      saveReceipt = True
'      lsSQL = "INSERT INTO Receipt_Master" & _
'             " SET sTransNox = " & strParm(GetNextCode("Receipt_Master", "sTransNox", True, p_oAppDrivr.Connection, True, p_oAppDrivr.BranchCode)) & _
'                ", dTransact = " & dateParm(p_oMaster("dSalesxxx")) & _
'                ", sClientID = " & strParm(p_oMaster("sClientID")) & _
'                ", sSerialID = " & strParm(p_oMaster("sSerialID")) & _
'                ", sORNoxxxx = " & strParm(loFrm.ORNo) & _
'                ", sSystemCd = 'MC'" & _
'                ", sReferNox = " & strParm(p_oMaster("sSourceNo")) & _
'                ", sSourceCd = 'MCSl'" & _
'                ", cTranType = '5'" & _
'                ", sInsTypID = '0109'" & _
'                ", cRegisFrm = '6'" & _
'                ", sAcctNmbr = ''" & _
'                ", sCoCltID1 = ''" & _
'                ", sCoCltID2 = ''" & _
'                ", nTranTotl = " & p_oMaster("nUpSalesx") & _
'                ", nTranAmtx = " & p_oMaster("nUpSalesx") & _
'                ", nDiscount = 0" & _
'                ", nOthersxx = 0" & _
'                ", cTranStat = '0'" & _
'                ", sRemarksx = 'MAPFRE'" & _
'                ", dModified = " & dateParm(p_oAppDrivr.ServerDate) & _
'                ", sModified = " & strParm(p_oAppDrivr.UserID)
'      p_oAppDrivr.Execute lsSQL, "Receipt_Master"
'
'      saveReceipt = True
   End If

endProc:
   Exit Function
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
        .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
    End With
    With Err
        .Raise .Number, .Source, .Description
    End With
End Sub
Function CancelTransaction() As Boolean
   'she 2016-05-16
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "CancelTransaction()"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   
   
   'Test if record is posted
   If p_oMaster("cTranStat") = xeStatePosted Or _
      p_oMaster("cTranStat") = xeStateCancelled Then
      MsgBox "Record was posted/cancelled. You are no longer allowed to cancel the record!!!" & vbCrLf & vbCrLf & _
            "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
        
   lsSQL = "UPDATE MC_Accident_Insurance SET" & _
               " cTranStat = " & strParm(xeStateCancelled) & _
           " WHERE sTransNox = " & strParm(p_oMaster("sTransNox"))
           
   If p_oAppDrivr.Execute(lsSQL, "MC_Accident_Insurance", p_sBranchCd) = 0 Then
      MsgBox "Unable to cancel " & p_oMaster("sTransNox") & " from Bank Account Adjustment Table." & vbCrLf & _
            "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   CancelTransaction = True
   
endProc:
    Exit Function
errProc:
   ShowError lsOldProc
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

