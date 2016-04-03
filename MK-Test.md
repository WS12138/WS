# WS
#MK-test is my first program in the github,of course,I have wroten some code by Visual Basic.NET.
#This version is a VBA function,this is the begining for me,I will copy the code what I have written in this website.

'-----------------------------------------------------------------------------------------------------------------------'
'MK检验，时间序列设定为100个元素。
'对这个函数的可靠性需要修改。
Private sngData(100) As Single
Private sngDataOpposite(100) As Single

Private Sub btnMKTest_Click()
    
    Dim readData As Range            '存选中的单元格数据。
    Dim writeData As Range           '设定写数据的单元格位置。
    Dim newRange As Range
    
    Dim writeRangeX As Single        '单元格位置横坐标。
    Dim writeRangeY As Single        '单元格位置纵坐标。
    
    Dim iStep As Integer
    Dim i As Long
    Dim iTemp As Integer              '逆序序列用。
        
    Dim iSk As Integer
    Dim sngESk As Single
    Dim sngVarSk As Double

    Dim sngUF(0 To 100) As Single
    Dim sngUB(0 To 100) As Single

    Set readData = Application.InputBox("Select the ranges.", "Selection", Type:=8)
    
    '数据输入。
    '顺序序列，sngData数组从1开始。
    iStep = 0
    With readData
        For Each c In readData
            iStep = iStep + 1
            sngData(iStep) = c.Value
        Next
    End With
    
    '逆序序列。
    iTemp = iStep + 1
    For i = iStep To 1 Step -1
        sngDataOpposite(i) = sngData(iTemp - i)
    Next i

    sngUF(1) = 0
    For i = 2 To iStep
        iSk = iSk + Ri(i)
        sngESk = i * (i - 1) / 4
        sngVarSk = (i * (i - 1) * (2 * i + 5)) / 72      '(i * (i - 1) * (2 * i + 5))当i大于30时数据会溢出，将i设为long型数据就可以解决，原因？
        sngUF(i) = (iSk - sngESk) / (sngVarSk ^ 0.5)
    Next

    iSk = 0
    sngESk = 0
    sngVarSk = 0
    sngUB(1) = 0
    For i = 2 To iStep
        iSk = iSk + RiOpposite(i)
        sngESk = i * (i - 1) / 4
        sngVarSk = (i * (i - 1) * (2 * i + 5)) / 72
        sngUB(i) = (iSk - sngESk) / (sngVarSk ^ 0.5)
    Next
    
    Set writeData = Application.InputBox("Select the ranges.", "Selection", Type:=8)
    Set newRange = writeData
    'Range属性，确定选定区域的位置，但只是第一个单元格的属性。
    writeRangeY = writeData.Column
    writeRangeX = writeData.Row
    For i = 1 To iStep
         
        '输出UF。(Cell属性是相对值，是基于选定点的偏移量,而且Cell里的参数不能为0,当参数为1的时候即表示所在的行或列.)
        newRange.Cells(i, 1).Value = sngUF(i)
        '输出UB。
        newRange.Cells(i, 2).Value = sngUB(iTemp - i)
        
    Next
        
    MsgBox "Successful."
    
End Sub

'顺序用,确定正向累积数。
Private Function Ri(ByVal index As Integer) As Integer

    Dim iResult As Integer
    Dim i As Integer

    For i = 1 To (index - 1)
        If sngData(index) > sngData(i) Then
            iResult = iResult + 1
        End If
    Next i
    Ri = iResult

End Function

'逆序用，确定逆向累计数。
Private Function RiOpposite(ByVal index As Integer) As Integer

    Dim iResult As Integer
    Dim i As Integer

    For i = 1 To (index - 1)
        If sngDataOpposite(index) > sngDataOpposite(i) Then
            iResult = iResult + 1
        End If
    Next i
    RiOpposite = iResult

End Function
'According to the result of the test in the Excel2013,these codes is OK.
'-----------------------------------------------------------------------------------------------------------------------------'
