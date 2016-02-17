VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPInternalOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function create(pTestRun As String, pData As SAPOrderList)
    On Error GoTo create_Error
    Dim lSAPFormat As New SAPFormat
    Dim boInternalOrder As Object
    Dim oReturn As Object
    Dim oI_Master_Data As Object
    Dim oSrules As Object
    Dim oSruleRow As Object
    Dim oLanguageKey As Object
    Dim aControllingArea As String
    Dim aTestRun As String

    Set boInternalOrder = MySAPCon.SAPBapiControl.GetSAPObject("InternalOrder")
    Set oI_Master_Data = MySAPCon.SAPBapiControl.DimAs(boInternalOrder, "create", "I_Master_Data")
    Set oSrules = MySAPCon.SAPBapiControl.DimAs(boInternalOrder, "create", "Srules")

    oI_Master_Data("ORDER") = lSAPFormat.unpack(pData.ORDER, 12)
    oI_Master_Data.Value("ORDER_TYPE") = pData.ORDER_TYPE
    oI_Master_Data.Value("ORDER_NAME") = pData.ORDER_NAME
    oI_Master_Data.Value("CO_AREA") = pData.CO_AREA
    oI_Master_Data.Value("COMP_CODE") = pData.COMP_CODE
    oI_Master_Data.Value("BUS_AREA") = pData.BUS_AREA
    oI_Master_Data.Value("PLANT") = pData.PLANT
    oI_Master_Data.Value("PROFIT_CTR") = pData.PROFIT_CTR
    oI_Master_Data.Value("FUNC_AREA") = pData.FUNC_AREA
    If pData.GL_ACCOUNT <> "" Then
        Set oSruleRow = oSrules.Rows.Add
        oSruleRow.Value("SETTL_TYPE") = "FUL"
        oSruleRow.Value("COMP_CODE") = pData.COMP_CODE
        oSruleRow.Value("PERCENTAGE") = 100
        oSruleRow.Value("GL_ACCOUNT") = lSAPFormat.unpack(pData.GL_ACCOUNT, 10)
    End If
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boInternalOrder.create I_Master_Data:=oI_Master_Data, _
        Srules:=oSrules, _
        TestRun:=pTestRun, _
        Return:=oReturn
    Else
        boInternalOrder.create I_Master_Data:=oI_Master_Data, _
        Srules:=oSrules, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit
    Dim aErr As Integer
    If oReturn.RowCount = 0 Then
        create = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        aErr = 0
        For Each retRow In oReturn.Rows
            create = create & ";" & retRow("MESSAGE")
            If retRow("TYPE") = "E" Then
                aErr = 1
            End If
        Next retRow
        If aErr = 0 Then
            create = "Success: " & create
            aSAPBapiTranctionCommit.commit
        End If
    End If
    Set boInternalOrder = Nothing
    Exit Function

create_Error:
    Application.Cursor = xlDefault
    create = "Error"
    MySAPErr.MSGProt "SAPInternalOrder", "create", "", Err.Number, Err.Description
    Exit Function
End Function

