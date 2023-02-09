VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call EvaluateRow
End Sub

Private Sub CommandButton2_Click()
    Call EvaluateAll
End Sub

Private Sub CommandButton3_Click()
    Call Report
End Sub
