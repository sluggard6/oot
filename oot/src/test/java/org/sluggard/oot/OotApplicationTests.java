package org.sluggard.oot;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.sluggard.oot.bean.TableInfo;
import org.sluggard.oot.dao.SimpleDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class OotApplicationTests {
	
	@Autowired
	private SimpleDao simpleDao;
	
	private static Pattern TABLE_NAME = Pattern.compile("^[A-Za-z_]+$");

	@Test
	void contextLoads() throws Exception {
		List<TableInfo> list = simpleDao.runSimpleSql();
		Map<String, List<TableInfo>> map = new TreeMap<>();
		list.forEach(t -> {
			System.out.println(t);
			if(!map.containsKey(t.getTableName())) {
				List<TableInfo> table = new ArrayList<>();
				map.put(t.getTableName(), table);
			}
			map.get(t.getTableName()).add(t);
		});
//		writeToExcel(map, "aa.xlsx");
		writeToWord(map, "aa.docx");;
	}
	
	private void writeToWord(Map<String, List<TableInfo>> map, String fileName) throws Exception {
		XWPFDocument doc = new XWPFDocument();
		addCustomHeadingStyle(doc, "1", 1);
    	XWPFParagraph paragraph = doc.createParagraph();
    	paragraph.setPageBreak(true);
    	paragraph.setAlignment(ParagraphAlignment.CENTER);
    	XWPFRun run = paragraph.createRun();
    	run.setFontSize(40);
    	run.setText("宝龙PD系统数据字典");
		map.forEach((k, v)->{
			if(TABLE_NAME.matcher(k).matches()){
				createTableParagraph(doc, k, v);
			}
		});
		File file = new File(fileName);
		if(file.exists()) {
			file.delete();
			file.createNewFile();
		}
		FileOutputStream stream = new FileOutputStream(file);
        // 需要抛异常
		doc.write(stream);
         //关流
        stream.close();
	}
	
	/**
     * 增加自定义标题样式。这里用的是stackoverflow的源码
     * 
     * @param docxDocument 目标文档
     * @param strStyleId 样式名称
     * @param headingLevel 样式级别
     */
    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {
 
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);
 
        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);
 
        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));
 
        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);
 
        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);
 
        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);
 
        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);
 
        XWPFStyle style = new XWPFStyle(ctStyle);
 
        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();
 
        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
 
    }
	
	/**
     * 设置表格
     * @param document
     * @param rows
     * @param cols
     * @Author Huangxiaocong 2018年12月16日 
     */
    public void createTableParagraph(XWPFDocument document, String tableName, List<TableInfo> map) {
//        xwpfHelperTable.createTable(xdoc, rowSize, cellSize, isSetColWidth, colWidths)
    	XWPFParagraph paragraph = document.createParagraph();
    	paragraph.setStyle("1");
    	paragraph.setPageBreak(true);
    	XWPFRun run = paragraph.createRun();
    	String tComments = map.get(0).gettComments();
    	if(StringUtils.isNoneBlank(tComments)) {
    		run.setText(String.format("%s : %s", tableName, map.get(0).gettComments()));
    	}else {
    		run.setText(tableName);
    	}
    	
    	
    	XWPFTable table = document.createTable(map.size()+1, 6);
    	fillTableData(table, tableName, map);
    }
    
    /**
     * 往表格中填充数据
     * @param table
     * @param tableData
     * @Author Huangxiaocong 2018年12月16日
     */
    public void fillTableData(XWPFTable table, String tableName, List<TableInfo> tableData) {
        List<XWPFTableRow> rowList = table.getRows();
        List<XWPFTableCell> cellList = rowList.get(0).getTableCells();
        cellList.get(0).setText("列名");
        cellList.get(1).setText("数据类型");
        cellList.get(2).setText("长度");
        cellList.get(3).setText("小数位");
        cellList.get(4).setText("默认值");
        cellList.get(5).setText("说明");
        for(int i = 0; i < tableData.size(); i++) {
			TableInfo ti = tableData.get(i);
            cellList = rowList.get(i+1).getTableCells();
            cellList.get(0).setText(ti.getColumnName());
            cellList.get(1).setText(ti.getDataType());
            cellList.get(2).setText(ti.getDataLength());
            cellList.get(3).setText(ti.getDataScale());
            cellList.get(4).setText(ti.getDataDefault());
            cellList.get(5).setText(ti.getcComments());
        }
    }
	
	private void writeToExcel(Map<String, List<TableInfo>> map, String fileName) throws Exception {
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFCellStyle style = workbook.createCellStyle();
		style.setBorderBottom(BorderStyle.THIN); //下边框
		style.setBorderLeft(BorderStyle.THIN);//左边框
		style.setBorderTop(BorderStyle.THIN);//上边框
		style.setBorderRight(BorderStyle.THIN);//右边框
		map.forEach((k, v)->{
			if(TABLE_NAME.matcher(k).matches()){
				XSSFSheet sheet = workbook.createSheet(k);
				XSSFRow row = sheet.createRow(0);
				XSSFCell cell = row.createCell(0);
				cell.setCellStyle(style);
				cell.setCellValue(k);
				cell = row.createCell(1);
				cell.setCellStyle(style);
				cell = row.createCell(2);
				cell.setCellStyle(style);
				cell = row.createCell(3);
				cell.setCellStyle(style);
				cell.setCellValue(v.get(0).gettComments());
				cell = row.createCell(4);
				cell.setCellStyle(style);
				cell = row.createCell(5);
				cell.setCellStyle(style);
//				row.createCell(3).setCellValue(v.get(0).gettComments());
				// 合并日期占两行(4个参数，分别为起始行，结束行，起始列，结束列)
		        // 行和列都是从0开始计数，且起始结束都会合并
		        // 这里是合并excel中日期的两行为一行
		        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
		        sheet.addMergedRegion(region);
		        region = new CellRangeAddress(0, 0, 3, 5);
		        sheet.addMergedRegion(region);
//		        XSSFCell cell;
				row = sheet.createRow(1);
				cell = row.createCell(0);
				cell.setCellStyle(style);
				cell.setCellValue("列名");
				cell = row.createCell(1);
				cell.setCellStyle(style);
				cell.setCellValue("数据类型");
				cell = row.createCell(2);
				cell.setCellStyle(style);
				cell.setCellValue("长度");
				cell = row.createCell(3);
				cell.setCellStyle(style);
				cell.setCellValue("小数位");
				cell = row.createCell(4);
				cell.setCellStyle(style);
				cell.setCellValue("默认值");
				cell = row.createCell(5);
				cell.setCellStyle(style);
				cell.setCellValue("说明");
				for(int i=0;i<v.size();i++) {
					row = sheet.createRow(i+2);
					TableInfo ti = v.get(i);
					cell = row.createCell(0);
					cell.setCellStyle(style);
					cell.setCellValue(ti.getColumnName());
					cell = row.createCell(1);
					cell.setCellStyle(style);
					cell.setCellValue(ti.getDataType());
					cell = row.createCell(2);
					cell.setCellStyle(style);
					cell.setCellValue(ti.getDataLength());
					cell = row.createCell(3);
					cell.setCellStyle(style);
					cell.setCellValue(ti.getDataScale());
					cell = row.createCell(4);
					cell.setCellStyle(style);
					cell.setCellValue(ti.getDataDefault());
					cell = row.createCell(5);
					cell.setCellStyle(style);
					cell.setCellValue(ti.getcComments());
				}
//				sheet.autoSizeColumn(0); //调整第一列宽度
//		        sheet.autoSizeColumn(1); //调整第二列宽度
//		        sheet.autoSizeColumn(2); //调整第三列宽度
//		        sheet.autoSizeColumn(3); //调整第四列宽度
//		        sheet.autoSizeColumn(4); //调整第四列宽度
//		        sheet.autoSizeColumn(5); //调整第四列宽度
				setSizeColumn(sheet, 6);
				sheet.autoSizeColumn(0); //调整第一列宽度
		        sheet.autoSizeColumn(1); //调整第二列宽度
			}
		});
		File file = new File(fileName);
		if(file.exists()) {
			file.delete();
			file.createNewFile();
		}
		FileOutputStream stream = new FileOutputStream(file);
        // 需要抛异常
        workbook.write(stream);
         //关流
        stream.close();
        System.out.println(workbook.getNumberOfSheets());
	}
	
	private void setSizeColumn(XSSFSheet sheet, int size) {
        for (int columnNum = 0; columnNum < size; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 1; rowNum < sheet.getLastRowNum(); rowNum++) {
                XSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
 
                if (currentRow.getCell(columnNum) != null) {
                    XSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == CellType.STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if(length>255) {
                        	columnWidth = 255;
                        }else if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
    }

}

//class Table {
//	private String tableName;
//	private String commnets;
//}
//
//class Column {
//	private String columnName;
//	private String dataType;
//	private String data
//}