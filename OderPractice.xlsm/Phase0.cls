VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Phase0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPhase

Public Sub IPhase_ExecutePhase()
    Debug.Print "phase:START_0"
        
    RefleshTable
    Message.ClearAll
    
    Dim record As SalesData
    Set record = Main.GetIndicator.GetRecord
    
    Dim ws As Worksheet
    Set ws = Worksheets("main")
    ws.Cells(MAIN_ROW.CARRY_OVER_STOCK, MAIN_COLUMN.Number) = record.GetStock
    ws.Cells(MAIN_ROW.DATE_, MAIN_COLUMN.Number) = record.GetDate
    ws.Cells(MAIN_ROW.CURRENT_STOCK, MAIN_COLUMN.Number) = record.GetStock
    
    ws.Cells(MAIN_ROW.PRICE_CUT, MAIN_COLUMN.Number) = Main.GetCabinet.PriceCutItems
    
  
End Sub

Public Sub IPhase_ChangePhase()
    Main.SetCurrentPhase (PhaseNumber.DELIVERY_1)
End Sub

'mainシートの当日の数字を消す。
Private Sub RefleshTable()
        Dim ws As Worksheet
        Set ws = Worksheets("main")
        
        Dim i As Integer
        For i = 2 To 7
            ws.Cells(i, MAIN_COLUMN.Number).ClearContents
        Next i
        For i = 2 To 6
            ws.Cells(i, MAIN_COLUMN.Price).ClearContents
        Next i
        
End Sub
