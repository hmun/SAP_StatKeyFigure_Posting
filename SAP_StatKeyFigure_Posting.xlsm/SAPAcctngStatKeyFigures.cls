VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPAcctngStatKeyFigures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsConnected As Integer

Private Sub Class_Initialize()
    IsConnected = SAPCheck()
End Sub

Public Function post(pKokrs As String, pBuDat As String, pBldat As String, pData As Collection) As String
    On Error GoTo post_Error
    Dim boAcctngStatKeyFigures As Object
    Dim oDocHeader As Object
    Dim oDocItems As Object
    Dim oReturn As Object
    Dim aDocNo As String
    Dim aIgnoreWarnings As String
    Dim lSAPFormat As New SAPFormat
    Dim lSubrc As Boolean

    Set boAcctngStatKeyFigures = MySAPCon.SAPBapiControl.GetSAPObject("AcctngStatKeyFigures")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boAcctngStatKeyFigures, "Post", "Return")
    Set oDocHeader = MySAPCon.SAPBapiControl.DimAs(boAcctngStatKeyFigures, "Post", "DocHeader")
    Set oDocItems = MySAPCon.SAPBapiControl.DimAs(boAcctngStatKeyFigures, "Post", "DocItems")
    oDocHeader("CO_AREA") = pKokrs
    oDocHeader("DOCDATE") = pBldat
    oDocHeader("POSTGDATE") = pBuDat
    oDocHeader("USERNAME") = MySAPCon.SAPCon.User
    Dim aItemRow As Object
    Dim lRow As Object
    For Each lRow In pData
        Set aItemRow = oDocItems.Rows.Add
        aItemRow.Value("STATKEYFIG") = lSAPFormat.unpack(lRow.STATKEYFIG, 6)
        aItemRow.Value("STAT_QTY") = Format$(lRow.STAT_QTY, "0.00")
        aItemRow.Value("SEG_TEXT") = lRow.SEG_TEXT
        aItemRow.Value("REC_CCTR") = lSAPFormat.unpack(lRow.REC_CCTR, 10)
        aItemRow.Value("ACTTYPE") = lRow.ACTTYPE
    Next lRow

    boAcctngStatKeyFigures.post DocHeader:=oDocHeader, _
    IgnoreWarnings:="X", _
    DocItems:=oDocItems, _
    DocNo:=oDocNo, _
    Return:=oReturn
    Set aRow = oReturn.Rows(1)
    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    lSubrc = True
    If aRow("TYPE") = "S" Then
        post = ";" & aRow("MESSAGE")
    Else
        Dim RetRow As Object
        For Each RetRow In oReturn.Rows
            If RetRow("ID") <> "BK" Or RetRow("NUMBER") <> "057" Then
                post = post & ";" & RetRow("MESSAGE")
            End If
            If RetRow("TYPE") = "E" Or RetRow("TYPE") = "A" Then
                lSubrc = False
            End If
        Next RetRow
    End If
    If lSubrc = True Then
        aSAPBapiTranctionCommit.commit
    End If

    Set boAcctngStatKeyFigures = Nothing
    Exit Function
post_Error:
    post = "Fehler"
    MySAPErr.MSGProt "SAPAcctngStatKeyFigures", "post", "", Err.Number, Err.Description
    Exit Function
End Function
