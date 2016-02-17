VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ORDER As String
Public ORDER_TYPE As String
Public ORDER_NAME As String
Public CO_AREA As String
Public COMP_CODE As String
Public BUS_AREA As String
Public PLANT As String
Public PROFIT_CTR As String
Public FUNC_AREA As String
Public GL_ACCOUNT As String

Public Function create(p_ORDER, p_ORDER_TYPE, p_ORDER_NAME, p_CO_AREA, p_COMP_CODE, p_BUS_AREA, p_PLANT, p_PROFIT_CTR, p_FUNC_AREA, p_GL_ACCOUNT)
    ORDER = p_ORDER
    ORDER_TYPE = p_ORDER_TYPE
    ORDER_NAME = p_ORDER_NAME
    CO_AREA = p_CO_AREA
    COMP_CODE = p_COMP_CODE
    BUS_AREA = p_BUS_AREA
    PLANT = p_PLANT
    PROFIT_CTR = p_PROFIT_CTR
    FUNC_AREA = p_FUNC_AREA
    GL_ACCOUNT = p_GL_ACCOUNT
End Function
