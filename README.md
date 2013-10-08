
##fluentexcel

>简介

> fluentexcel基于POI,旨在改进 excel 生成时的编码体验, 借鉴 fluent api 思想开发方便易用，同时书写方便的 excel 的 api。 更重要的是 fluentexcel 可以将一个 excel 2007 的文档直接生成为 excel 2003 格式， 程序可以只关注处理逻辑， 忽略样式差异。

> fluentexcel 目前包含的内容、以及支持的功能都是源于工作中对excel的处理，所以目前未能完全覆盖POI中提供的所有功能

###如何使用fluentexcel?

fluentexcel 为什么会存在, 这得先看看 单纯使用原生 POI 如何产生一个excel

这是一个 excel 2003 的例子, 

    /**
	 *
	 * 单元格颜色填充
	 */
	public static void fillColorTest(){
		String file = "res/fillColortest.xls";
		Workbook wb = new HSSFWorkbook();
	    Sheet sheet = wb.createSheet("new sheet");

	    // Create a row and put some cells in it. Rows are 0 based.
	    Row row = sheet.createRow((short) 1);

	    // Aqua background
	    CellStyle style = wb.createCellStyle();
	    style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    Cell cell = row.createCell((short) 1);
	    cell.setCellValue("X");
	    cell.setCellStyle(style);

	    // Orange "foreground", foreground being the fill foreground not the font color.
	    style = wb.createCellStyle();
	    style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    cell = row.createCell((short) 2);
	    cell.setCellValue("X");
	    cell.setCellStyle(style);

	    FileOutputStream fileout;
		try {
			fileout = new FileOutputStream(file);
			wb.write(fileout);
			fileout.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

然后再看 fluentexcel 如何实现相同的功能

    Render render = new XlsxRender();
    ExcelBuilder excelBuilder = new ExcelBuilder();

    /* 使用ExcelBuilder 创建样式 */
    CellStyle styleBlack = excelBuilder
            .createCellStyle()
                .withRender(render)
                .useFgColor((short) 0, (short) 0, (short) 0)
                .createFont()
                    .boldWeight((short)8).fontName("微软雅黑").italic().size((short)14)
                .fontOver()
            .cellStyleOver();
    
    CellStyle styleRed = excelBuilder
            .createCellStyle()
                .withRender(render)
                .useFgColor((short) 254, (short) 0, (short) 0)
                .createFont()
                    .boldWeight((short)8).fontName("微软雅黑").italic().size((short)14)
                .fontOver()
            .cellStyleOver();
    
    /* 使用ExcelBuilder 创建文档 */
    excelBuilder
            .createDoc().docName("fluentTest2").type("xlsx").withDocRender(render)
                .createSheet().sheetName("Hello")
                    .createBlock().fromRow(2)
                        .createRow().fromCol(1)
                            .createCell().content("X").withStyle(styleBlack).cellOver()
                            .createCell().content("X").withStyle(styleRed).cellOver()
                        .rowOver()
                    .blockOver()
                .sheetOver()
            .saveTo("d:/app/tmp/excel")
            .rendDoc(); 


###还有什么功能?


正如上面的示例，你可以通过流式的方法链创建excel和其样式。fluentexcel 还提供了一些其他的工具方便创建excel文档

**单元格样式定义**


不过如果你觉得 fluentexcel 创建样式的方式看起来不是那么清晰，那么 fluentexcel 也提供了通过 json 定义单元格样式的方法。

用法如下：

首先定义出样式
    
    String style = "{'font':{'size':8, 'name':'微软雅黑','boldWeight':8},'fgColor':{'r':255, 'g':192, 'b':0}}"
    
然后使用Render 将 json格式的style转化为 CellStyle

    CellStyle cellStyle = render.parse(wb, style);
    
    
**excel 2003** 和 **excel 2007**

除了样式之外，如果要生成 excel 2003 的excel 同样也非常简便。 对于一个电子表格来说, 如果要生成不同版本的excel, 数据是抽象的，因此是不同的，所以在支持多种excel格式上, fluentexcel 将POI提供的 `CellStyle` 定义再封装一遍形成 `Render`, excel 2003 和 excel 2007 分别为 `XlsRender`, `XlsxRender`。 用户不需要再关注产生电子表格时类库细微的差别。

使用方法：

    Render xlsxRender = new XlsxRender();  // excel 2007
    Render xlsRender = new XlsRender();    // excel 2003
    
之后处理excel 的方法完全相同。 老板再也不用担心多版本excel支持，有木有！！！
    
###还有什么需要改进?

1. 对图片的支持
2. Poi中所有的CellStyle 还没有完全支持
3. ... 还有很多，这里不一一列举






