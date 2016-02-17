Attribute VB_Name = "SAPMakro"
Sub SAP_StatKeyFig_post()
    Dim aSAPAcctngStatKeyFigures As New SAPAcctngStatKeyFigures
    Dim aSAPDocItem As New SAPDocItem
    Dim aData As New Collection
    Dim aRetStr As String
    Dim aDateFormatString As New DateFormatString

    Dim aKOKRS As String
    Dim aMAXLINES As Integer
    Dim aFromLine As Integer
    Dim aToLine As Integer

    Dim aBLDAT As String
    Dim aBUDAT As String
    Dim aMENGE As String
    Dim aEPSP As String
    Dim aSKOSTL As String
    Dim aLEART As String

    Worksheets("Parameter").Activate
    aBUDAT = Format(Cells(2, 2), aDateFormatString.getString)
    aBLDAT = Format(Cells(3, 2), aDateFormatString.getString)
    aKOKRS = Format(Cells(4, 2), "0000")
    aMAXLINES = CInt(Cells(5, 2).Value)

    If IsNull(aBUDAT) Or aBUDAT = "" Or _
        IsNull(aBLDAT) Or aBLDAT = "" Or _
        IsNull(aKOKRS) Or aKOKRS = "" Or _
        IsNull(aMAXLINES) Or aMAXLINES = 0 Then
        MsgBox "Fill all mandatory fields in sheet Parameter!", vbCritical + vbOKOnly
        Exit Sub
    End If
    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    Worksheets("Data").Activate
    i = 2

    Dim aPostingLine As Integer
    aPostingLine = 1
    Do
        If Cells(i, 4) <> 0 And Left(Cells(i, 5).Value, 5) <> ";Docu" Then
            Set aSAPDocItem = New SAPDocItem
            aSAPDocItem.create Cells(i, 3).Value, CDbl(Cells(i, 5).Value), _
            Cells(i, 4).Value, Cells(i, 1).Value, Cells(i, 2).Value
            aData.Add aSAPDocItem
            If aPostingLine >= aMAXLINES Then
                aRetStr = aSAPAcctngStatKeyFigures.post(aKOKRS, aBUDAT, aBLDAT, aData)
                For j = 1 To aData.Count
                    Cells(i - (j - 1), 6) = aRetStr
                Next j
                Set aData = New Collection
                aPostingLine = 1
                aDatCompare = aDatCurrent
            End If
        End If
        i = i + 1
    Loop While Not IsEmpty(Cells(i, 1).Value)
    ' post the rest
    If aData.Count > 0 Then
        aRetStr = aSAPAcctngStatKeyFigures.post(aKOKRS, aBUDAT, aBLDAT, aData)
        Cells(i - 1, 6) = aRetStr
    End If
End Sub
