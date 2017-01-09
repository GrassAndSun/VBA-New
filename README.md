# VBA-New
Sub 按钮1_Click()
'定义交易基本信息
Dim Initial_Capital As Double '定义初始资金
Dim ETF_Price As Double '定义ETF价格
Dim Future_Price As Double '定义期货价格
'定义基本面限制条件
Dim Max_Future_No As Integer '定义期货单日最高交易数量
Dim Margin_Ratio As Double '定义保证金比例
Dim Margin_Multiplier As Double '定义保证金弹性倍数
Dim Contract_Multiplier As Integer '定义合约乘数
'定义输出结果
Dim ETF_No As Double '定义ETF数量
Dim Future_No As Integer '定义期货数量
Dim Cash As Double '定义剩余现金
Dim Hedge_Ratio As Double '定义对冲比例
Dim Utilization_Efficiency_of_Captial As Double '定义资金使用率
'定义临时变量
Dim Half_Initial_Capital As Double '定义初始资金一半
Dim i As Integer


'加载交易基本信息
Initial_Capital = Sheet2.Cells(2, 2) '初始资金赋值
ETF_Price = Sheet2.Cells(5, 2) 'ETF价格赋值
Future_Price = Sheet2.Cells(6, 2) '期货价格赋值

'加载基本面限制条件
Max_Future_No = Sheet1.Cells(2, 2) '单日最大交易赋值
Margin_Ratio = Application.WorksheetFunction.VLookup(Sheet2.Cells(4, 2), Sheet1.Range("B3:C4"), 2, False) '保证金赋值
Margin_Multiplier = Sheet1.Cells(5, 2) '保证金弹性倍数赋值
Contract_Multiplier = Application.WorksheetFunction.VLookup(Sheet2.Cells(3, 2), Sheet1.Range("B6:C8"), 2, False) '合约乘数赋值

'开始计算
Half_Initial_Capital = Initial_Capital / 2 '计算资金一半
Future_No = Fix(Half_Initial_Capital / (Future_Price * Contract_Multiplier)) '计算应该购买期货的数量
 '限定期货单日最高交易数量
If Future_No > Max_Future_No Then
    Future_No = Max_Future_No
End If
ETF_No = Fix((Future_No * Future_Price * Contract_Multiplier) / ETF_Price / 100) * 100 '计算ETF的购买数量
Hedge_Ratio = ETF_No * ETF_Price / (Future_Price * Future_No * Contract_Multiplier) '计算对冲比例

'提高对冲比例
While Hedge_Ratio < 0.9
Future_No = Future_No - 1
ETF_No = Fix((Future_No * Future_Price * Contract_Multiplier) / ETF_Price / 100) * 100
Hedge_Ratio = ETF_No * ETF_Price / (Future_Price * Future_No * Contract_Multiplier)
If Future_No = 0 Then
    MsgBox ("对冲资金不足")
    Exit Sub
End If
Wend


Cash = Initial_Capital - ETF_No * ETF_Price - (Future_Price * Contract_Multiplier * Margin_Ratio * Margin_Multiplier) '计算剩余资金
Utilization_Efficiency_of_Captial = (Initial_Capital - Cash) / Initial_Capital '计算资金使用比例

'输出结果
Sheet2.Cells(2, 9) = ETF_No
Sheet2.Cells(3, 9) = Future_No
Sheet2.Cells(4, 9) = Cash
Sheet2.Cells(5, 9) = Hedge_Ratio
Sheet2.Cells(6, 9) = Utilization_Efficiency_of_Captial

End Sub
 
