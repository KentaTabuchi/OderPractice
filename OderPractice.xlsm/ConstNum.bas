Attribute VB_Name = "ConstNum"
Option Explicit
'data�V�[�g�̗�萔
Public Const COLUMN_DATE = 2
Public Const COLUMN_DELIVERY = 3
Public Const COLUMN_SALES = 4
Public Const COLUMN_LOSS = 5
Public Const COLUMN_STOCK = 6

Enum DATA_COLUMN
    DATE_2 = 2
    DELIVERY_3 = 3
    SALES_4 = 4
    LOSS_5 = 5
    STOCK_6 = 6
    ODER_7 = 7
End Enum

'main�V�[�g�̍s�萔
Enum MAIN_ROW
    DATE_ = 2
    CARRY_OVER_STOCK = 3
    PRICE_CUT = 4
    DELIVERY = 5
    sales = 6
    loss = 7
    CURRENT_STOCK = 8
    ITEMS = 15
End Enum
'main�V�[�g�̗�萔
Enum MAIN_COLUMN
    Number = 7
    Price = 8
End Enum

Enum PhaseNumber
    START_0 = 0
    DELIVERY_1 = 1
    SELL_2 = 2
    LOSS_3 = 3
    LAST_4 = 4
End Enum




