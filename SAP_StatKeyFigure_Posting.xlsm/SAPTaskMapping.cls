VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPTaskMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ACTTYPE As String
Public REC_WBS_EL As String

Public Function create(pACTTYPE As String, pREC_WBS_EL As String)
    ACTTYPE = pACTTYPE
    REC_WBS_EL = pREC_WBS_EL
End Function

