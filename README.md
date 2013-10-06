
##fluentexcel

>���

> fluentexcel����POI,ּ�ڸĽ� excel ����ʱ�ı�������, ��� fluent api ˼�뿪���������ã�ͬʱ��д����� excel �� api�� ����Ҫ���� fluentexcel ���Խ�һ�� excel 2007 ���ĵ�ֱ������Ϊ excel 2003 ��ʽ�� �������ֻ��ע�����߼��� ������ʽ���졣

> fluentexcel Ŀǰ���������ݡ��Լ�֧�ֵĹ��ܶ���Դ�ڹ����ж�excel�Ĵ�������Ŀǰδ����ȫ����POI���ṩ�����й���

###���ʹ��fluentexcel?

fluentexcel Ϊʲô�����, ����ȿ��� ����ʹ��ԭ�� POI ��β���һ��excel

����һ�� excel 2003 ������, 

    /**
	 *
	 * ��Ԫ����ɫ���
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

Ȼ���ٿ� fluentexcel ���ʵ����ͬ�Ĺ���

    Render render = new XlsxRender();
    ExcelBuilder excelBuilder = new ExcelBuilder();

    /* ʹ��ExcelBuilder ������ʽ */
    CellStyle styleBlack = excelBuilder
            .createCellStyle()
                .withRender(render)
                .useFgColor((short) 0, (short) 0, (short) 0)
                .createFont()
                    .boldWeight((short)8).fontName("΢���ź�").italic().size((short)14)
                .fontOver()
            .cellStyleOver();
    
    CellStyle styleRed = excelBuilder
            .createCellStyle()
                .withRender(render)
                .useFgColor((short) 254, (short) 0, (short) 0)
                .createFont()
                    .boldWeight((short)8).fontName("΢���ź�").italic().size((short)14)
                .fontOver()
            .cellStyleOver();
    
    /* ʹ��ExcelBuilder �����ĵ� */
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


###����ʲô����?


���������ʾ���������ͨ����ʽ�ķ���������excel������ʽ��fluentexcel ���ṩ��һЩ�����Ĺ��߷��㴴��excel�ĵ�

**��Ԫ����ʽ����**


������������ fluentexcel ������ʽ�ķ�ʽ������������ô��������ô fluentexcel Ҳ�ṩ��ͨ�� json ���嵥Ԫ����ʽ�ķ�����

�÷����£�

���ȶ������ʽ
    
    String style = "{'font':{'size':8, 'name':'΢���ź�','boldWeight':8},'fgColor':{'r':255, 'g':192, 'b':0}}"
    
Ȼ��ʹ��Render �� json��ʽ��styleת��Ϊ CellStyle

    CellStyle cellStyle = render.parse(wb, style);
    
    
**excel 2003** �� **excel 2007**

������ʽ֮�⣬���Ҫ���� excel 2003 ��excel ͬ��Ҳ�ǳ���㡣 ����һ�����ӱ����˵, ���Ҫ���ɲ�ͬ�汾��excel, �����ǳ���ģ�����ǲ�ͬ�ģ�������֧�ֶ���excel��ʽ��, fluentexcel ��POI�ṩ�� `CellStyle` �����ٷ�װһ���γ� `Render`, excel 2003 �� excel 2007 �ֱ�Ϊ `XlsRender`, `XlsxRender`�� �û�����Ҫ�ٹ�ע�������ӱ��ʱ���ϸ΢�Ĳ��

ʹ�÷�����

    Render xlsxRender = new XlsxRender();  // excel 2007
    Render xlsRender = new XlsRender();    // excel 2003
    
֮����excel �ķ�����ȫ��ͬ�� �ϰ���Ҳ���õ��Ķ�汾excel֧�֣���ľ�У�����
    
###����ʲô��Ҫ�Ľ�?

1. ��ͼƬ��֧��
2. Poi�����е�CellStyle ��û����ȫ֧��
3. ... ���кܶ࣬���ﲻһһ�о�






