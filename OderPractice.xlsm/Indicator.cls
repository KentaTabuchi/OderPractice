VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Indicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_row As Integer
Private m_salesData As SalesData

Private Sub Class_Initialize()
    m_row = 2
    Set m_salesData = New SalesData
End Sub


Public Function GetRow() As Integer
    GetRow = m_row
End Function

Public Function MoveOn() As SalesData
    If (m_row = 32) Then
        m_row = 32
    Else
        m_row = m_row + 1
    End If
    PrintCursor
    Set MoveOn = GetRecord
End Function

Public Function TurnBack() As SalesData
    If (m_row = 2) Then
        m_row = 2
    Else
        m_row = m_row - 1
    End If
    PrintCursor
    Set TurnBack = GetRecord
End Function

Public Function GetRecord() As SalesData
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    
    m_salesData.SetDate (ws.Cells(m_row, COLUMN_DATE))
    m_salesData.SetDelivery (ws.Cells(m_row, COLUMN_DELIVERY))
    m_salesData.SetSales (ws.Cells(m_row, COLUMN_SALES))
    m_salesData.SetLoss (ws.Cells(m_row, COLUMN_LOSS))

    Set GetRecord = m_salesData
        
End Function

Private Sub PrintCursor()
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    ws.Columns(1).ClearContents
    ws.Cells(m_row, 1) = "Å®"
    
End Sub

'dataÉVÅ[ÉgÇ÷ÇÃèëÇ´çûÇ›èàóù
Public Sub UpdateRecord(Optional sales_ As Integer, Optional loss_ As Integer, Optional stock_ As Integer, Optional oder_ As Integer)
    
    Debug.Print "oder_:" & oder_
    If (sales_ <> 0) Then m_salesData.SetSales (sales_)
    If (loss_ <> 0) Then m_salesData.SetLoss (loss_)
    If (stock_ <> 0) Then m_salesData.SetStock (stock_)
    If (oder_ <> 0) Then m_salesData.SetOder (oder_)
     
    Dim ws As Worksheet
    Set ws = Worksheets("data")
    ws.Cells(m_row, DATA_COLUMN.SALES_4) = m_salesData.GetSales
    ws.Cells(m_row, DATA_COLUMN.LOSS_5) = m_salesData.GetLoss
    ws.Cells(m_row, DATA_COLUMN.STOCK_6) = m_salesData.GetStock
    ws.Cells(m_row, DATA_COLUMN.ODER_7) = m_salesData.GetOder
       
End Sub


