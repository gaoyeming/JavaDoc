'******************************************************************************
'* File:     Exported_Excel_page.vbs
'* Purpose:  分目录递归，查找当前PDM下所有表，并导出Excel
'* Title:    
'* Category: 
'* Version:  1.0
'* Author:  787681084@qq.com
'******************************************************************************

Option Explicit
ValidationMode = True
InteractiveMode = im_Batch

'-----------------------------------------------------------------------------
' 主函数
'-----------------------------------------------------------------------------
' 获取当前活动模型
Dim mdl ' 当前的模型
Set mdl = ActiveModel
Dim EXCEL,catalog,sheet,catalogNum,rowsNum,linkNum
rowsNum = 2
catalogNum = 1
linkNum = 1

If (mdl Is Nothing) Then
    MsgBox "当前模板不存在表"
Else
    SetCatalog
	ListObjects(mdl)
End If

'----------------------------------------------------------------------------------------------
' 设置目录索引sheet表头
'----------------------------------------------------------------------------------------------
Sub SetCatalog()
    Set EXCEL= CreateObject("Excel.Application")
    
    ' 使Excel通过应用程序对象可见。
    EXCEL.Visible = True
    EXCEL.workbooks.add(-4167)'添加工作表
	EXCEL.workbooks(1).sheets(1).name ="结尾sheet"
	EXCEL.workbooks(1).sheets.add
    EXCEL.workbooks(1).sheets(1).name ="目录索引"
    set catalog = EXCEL.workbooks(1).sheets("目录索引")
   
    catalog.cells(catalogNum, 1) = "序号"
    catalog.cells(catalogNum, 2) = "表名"
    catalog.cells(catalogNum, 3) = "表注释"
    
    ' 设置列宽和自动换行
    catalog.Columns(1).ColumnWidth = 6
    catalog.Columns(2).ColumnWidth = 30
    catalog.Columns(3).ColumnWidth = 60
    
    '设置首行居中显示
    catalog.Range(catalog.cells(1,1),catalog.cells(1,3)).HorizontalAlignment = 3
    '设置首行字体加粗
    catalog.Range(catalog.cells(1,1),catalog.cells(1,3)).Font.Bold = True
	 '设置边框   
    catalog.Range(catalog.cells(1, 1),catalog.cells(1, 3)).Borders.LineStyle = "1"  
    '设置背景颜色  
    catalog.Range(catalog.cells(1, 1),catalog.cells(1, 3)).Interior.ColorIndex = "19"  
End Sub

'----------------------------------------------------------------------------------------------
' 循环处理表
'----------------------------------------------------------------------------------------------
Private Sub ListObjects(fldr)
    output "开始循环生成 " + fldr.code + " 数据库中所有的表"
    Dim obj ' 运行对象
    For Each obj In fldr.children
        ' 调用子过程来打印对象上的信息
        DescribeObject obj
    Next
	output fldr.code + " 数据库中所有的表生成完成"
End Sub

'-----------------------------------------------------------------------------
' 表处理
'-----------------------------------------------------------------------------
Private Sub DescribeObject(CurrentObject)
    if not CurrentObject.Iskindof(cls_NamedObject) then exit sub
    if CurrentObject.Iskindof(cls_Table) then
	    ExportCatalog CurrentObject
		AddSheet CurrentObject.code
		ExportTable CurrentObject, sheet
		output CurrentObject.Name + "生成成功"
    End if
End Sub

'----------------------------------------------------------------------------------------------
' 导出目录结构
'----------------------------------------------------------------------------------------------
Sub ExportCatalog(tab)
    catalogNum = catalogNum + 1
    catalog.cells(catalogNum, 1).Value = catalogNum-1
    catalog.cells(catalogNum, 2).Value = tab.code
    catalog.cells(catalogNum, 3).Value = tab.comment
    '设置超链接
    catalog.Hyperlinks.Add catalog.cells(catalogNum,2), "",tab.code&"!A2"
	'设置序号居中显示
    catalog.Range(catalog.cells(catalogNum, 1),catalog.cells(catalogNum, 1)).HorizontalAlignment = 3
End Sub

'----------------------------------------------------------------------------------------------
' 新增sheet页;用于写入每张表需要的表头
'----------------------------------------------------------------------------------------------
Sub AddSheet(sheetName)
    EXCEL.workbooks(1).Sheets(2).Select
    EXCEL.workbooks(1).sheets.add
    EXCEL.workbooks(1).sheets(2).name = sheetName
    set sheet = EXCEL.workbooks(1).sheets(sheetName)
    '将一些文本放在工作表的第一行
	rowsNum = 2
    sheet.Cells(1, 1).Value = "目录索引"
	sheet.Cells(rowsNum, 1).Value = "序号"
    sheet.Cells(rowsNum, 2).Value = "字段名"
    sheet.Cells(rowsNum, 3).Value = "字段类型"
    sheet.Cells(rowsNum, 4).Value = "字段备注"
    sheet.cells(rowsNum, 5).Value = "主键"
    sheet.cells(rowsNum, 6).Value = "非空"
    sheet.cells(rowsNum, 7).Value = "默认值"
    '设置列宽
    sheet.Columns(1).ColumnWidth = 6
    sheet.Columns(2).ColumnWidth = 20
    sheet.Columns(3).ColumnWidth = 20
    sheet.Columns(4).ColumnWidth = 20
    sheet.Columns(5).ColumnWidth = 5
    sheet.Columns(6).ColumnWidth = 5
    sheet.Columns(7).ColumnWidth = 40

	'合并单元格
	sheet.Range(sheet.cells(1,1),sheet.cells(1,7)).Merge 
    '设置表头居中显示
    sheet.Range(sheet.cells(rowsNum,1),sheet.cells(rowsNum,7)).HorizontalAlignment = 3
    '设置表头字体加粗
    sheet.Range(sheet.cells(rowsNum,1),sheet.cells(rowsNum,7)).Font.Bold = True
	 '设置边框   
    sheet.Range(sheet.cells(rowsNum, 1),sheet.cells(rowsNum, 7)).Borders.LineStyle = "1"  
    '设置背景颜色  
    sheet.Range(sheet.cells(rowsNum, 1),sheet.cells(rowsNum, 7)).Interior.ColorIndex = "44"
    
    linkNum = linkNum + 1
    '设置超链接
    sheet.Hyperlinks.Add sheet.cells(1,1), "","目录索引"&"!B"&linkNum
End Sub

'----------------------------------------------------------------------------------------------
' 将表信息导入对应的sheet
'----------------------------------------------------------------------------------------------
Sub ExportTable(tab, sheet)
    Dim col ' 运行列
    Dim colsNum
    colsNum = 0
    for each col in tab.columns
        colsNum = colsNum + 1
        rowsNum = rowsNum + 1
        sheet.Cells(rowsNum, 1).Value = rowsNum-2
        sheet.Cells(rowsNum, 2).Value = col.code
        sheet.Cells(rowsNum, 3).Value = col.datatype
        sheet.Cells(rowsNum, 4).Value = col.comment
        
        If col.Primary = true Then
            sheet.cells(rowsNum, 5) = "Y" 
        Else
            sheet.cells(rowsNum, 5) = "N" 
        End If
        If col.Mandatory = true Then
            sheet.cells(rowsNum, 6) = "Y" 
        Else
            sheet.cells(rowsNum, 6) = "N" 
        End If
        
        sheet.cells(rowsNum, 7).Value = col.defaultvalue
        '设置居中显示
        sheet.cells(rowsNum,5).HorizontalAlignment = 3
        sheet.cells(rowsNum,6).HorizontalAlignment = 3
		sheet.cells(rowsNum,1).HorizontalAlignment = 3
    next
	' 记录索引
	Dim tableSql
	tableSql =  Split(tab.preview, vbcrlf)
	Dim sum,firstsite,lastsite
	sum =0
	for sum = 0 To UBound(tableSql)
		If mid(Trim(tableSql(sum)),1,7) = "primary" Then
			rowsNum = rowsNum + 1
			sheet.Cells(rowsNum, 1).Value = "主键"
			sheet.Cells(rowsNum, 2).Value = "primary"
			'设置字体加粗
			sheet.Range(sheet.cells(rowsNum,2),sheet.cells(rowsNum,3)).Font.Bold = True
			'合并单元格
			sheet.Range(sheet.cells(rowsNum,3),sheet.cells(rowsNum,7)).Merge 
			firstsite = instr(Trim(tableSql(sum)),"(")
			lastsite = instrRev(Trim(tableSql(sum)),")")
			sheet.Cells(rowsNum, 3).Value = mid(Trim(tableSql(sum)),firstsite+1,lastsite-firstsite-1)
		End If
		If mid(Trim(tableSql(sum)),1,3) = "key" Then
			rowsNum = rowsNum + 1
			sheet.Cells(rowsNum, 1).Value = "索引"
			sheet.Cells(rowsNum, 2).Value = "key"
			'设置字体加粗
			sheet.Range(sheet.cells(rowsNum,2),sheet.cells(rowsNum,3)).Font.Bold = True
			'合并单元格
			sheet.Range(sheet.cells(rowsNum,3),sheet.cells(rowsNum,7)).Merge 
			firstsite = instr(Trim(tableSql(sum)),"(")
			lastsite = instrRev(Trim(tableSql(sum)),")")
			sheet.Cells(rowsNum, 3).Value = mid(Trim(tableSql(sum)),firstsite+1,lastsite-firstsite-1)
		End If
	next
End Sub