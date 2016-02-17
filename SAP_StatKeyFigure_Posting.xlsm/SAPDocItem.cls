VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public STATKEYFIG As String
Public STAT_QTY As Double
Public SEG_TEXT As String
Public REC_CCTR As String
Public ACTTYPE As String

Public Function create(p_STATKEYFIG As String, p_STAT_QTY As Double, p_SEG_TEXT As String, _
    p_REC_CCTR As String, p_ACTTYPE As String)
    
    STATKEYFIG = p_STATKEYFIG
    STAT_QTY = p_STAT_QTY
    SEG_TEXT = p_SEG_TEXT
    REC_CCTR = p_REC_CCTR
    ACTTYPE = p_ACTTYPE
End Function

