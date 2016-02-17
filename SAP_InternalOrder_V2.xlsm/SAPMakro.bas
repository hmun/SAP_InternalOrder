Attribute VB_Name = "SAPMakro"
Sub SAP_InternalOrder_create()
    Dim aSAPInternalOrder As New SAPInternalOrder
    Dim aSAPOrderList As New SAPOrderList
    Dim aData As New Collection

    Dim aTestRun As String

    Dim i As Integer
    Dim aRetStr As String

    Worksheets("Parameter").Activate
    aTestRun = Cells(2, 2).Value

    aRet = SAPCheck()
    If Not aRet Then
        MsgBox "Connectio to SAP failed!", vbCritical + vbOKOnly
        Exit Sub
    End If
    ' Read the Data
    Worksheets("Data").Activate
    i = 2
    Do
        Set aSAPOrderList = New SAPOrderList
        aSAPOrderList.create Cells(i, 1).Value, Cells(i, 2).Value, Cells(i, 3).Value, Cells(i, 4).Value, _
        Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value, Cells(i, 8).Value, _
        Cells(i, 9).Value, Cells(i, 10).Value
        aRetStr = aSAPInternalOrder.create(aTestRun, aSAPOrderList)
        Cells(i, 11) = aRetStr
        i = i + 1
    Loop While Not IsNull(Cells(i, 1)) And Cells(i, 1) <> ""
End Sub

