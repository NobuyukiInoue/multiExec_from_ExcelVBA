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
    
    Select Case Range("E2").Value
    
        Case "Shell関数":
            multiExec_pidSearch_Shell Worksheets("実行コマンド列"), Range("A2").Value, Range("B2").Value, Range("C2").Value, Range("D2").Value
        
        Case "WSH(Run)":
            multiExec_pidSearch_WSH Worksheets("実行コマンド列"), Range("A2").Value, Range("B2").Value, Range("C2").Value, Range("D2").Value, Range("E2").Value
        
        Case "WSH(Exec)":
            multiExec_pidSearch_WSH Worksheets("実行コマンド列"), Range("A2").Value, Range("B2").Value, Range("C2").Value, Range("D2").Value, Range("E2").Value
        
    End Select

End Sub

Private Sub Worksheet_Activate()
    Dim WR As Range
    
    Set WR = Range("A4:D5")
    CommandButton1.Left = WR.Left
    CommandButton1.Top = WR.Top
    CommandButton1.Width = WR.Width
    CommandButton1.Height = WR.Height
    
End Sub
