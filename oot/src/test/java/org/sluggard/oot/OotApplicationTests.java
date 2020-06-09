package org.sluggard.oot;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

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
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
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
		Map<String, List<TableInfo>> map = new HashMap<>();
		list.forEach(t -> {
			System.out.println(t);
			if(!map.containsKey(t.getTableName())) {
				List<TableInfo> table = new ArrayList<>();
				map.put(t.getTableName(), table);
			}
			map.get(t.getTableName()).add(t);
		});
		writeToExcel(map, "aa.xlsx");
	}
	
	private void writeToWord(Map<String, List<TableInfo>> map, String fileName) throws Exception {
		XWPFDocument doc = new XWPFDocument();
	}
	
	/**
     * 设置表格
     * @param document
     * @param rows
     * @param cols
     * @Author Huangxiaocong 2018年12月16日 
     */
    public void createTableParagraph(XWPFDocument document, int rows, int cols) {
//        xwpfHelperTable.createTable(xdoc, rowSize, cellSize, isSetColWidth, colWidths)

//    	XWPFHelperTable xwpfHelperTable = new XWPFHelperTable();
//    	XWPFHelper xwpfHelper = new XWPFHelper();
        XWPFTable infoTable = document.createTable(rows, cols);
//        xwpfHelperTable.setTableWidthAndHAlign(infoTable, "9072", STJc.CENTER);
//        //合并表格
//        xwpfHelperTable.mergeCellsHorizontal(infoTable, 1, 1, 5);
//        xwpfHelperTable.mergeCellsVertically(infoTable, 0, 3, 6);
//        for(int col = 3; col < 7; col++) {
//            xwpfHelperTable.mergeCellsHorizontal(infoTable, col, 1, 5);
//        }
        //设置表格样式
        List<XWPFTableRow> rowList = infoTable.getRows();
        for(int i = 0; i < rowList.size(); i++) {
            XWPFTableRow infoTableRow = rowList.get(i);
            List<XWPFTableCell> cellList = infoTableRow.getTableCells();
            for(int j = 0; j < cellList.size(); j++) {
                XWPFParagraph cellParagraph = cellList.get(j).getParagraphArray(0);
                cellParagraph.setAlignment(ParagraphAlignment.CENTER);
                XWPFRun cellParagraphRun = cellParagraph.createRun();
                cellParagraphRun.setFontSize(12);
                if(i % 2 != 0) {
                    cellParagraphRun.setBold(true);
                }
            }
        }
//        xwpfHelperTable.setTableHeight(infoTable, 560, STVerticalJc.CENTER);
    }
    
    /**
     * 往表格中填充数据
     * @param table
     * @param tableData
     * @Author Huangxiaocong 2018年12月16日
     */
    public void fillTableData(XWPFTable table, List<TableInfo> tableData) {
        List<XWPFTableRow> rowList = table.getRows();
        List<XWPFTableCell> cellList = rowList.get(0).getTableCells();
        XWPFParagraph cellParagraph = cellList.get(0).getParagraphArray(0);
        XWPFRun cellParagraphRun = cellParagraph.getRuns().get(0);
        cellParagraph.getRuns().get(0).setText("列名");
        cellParagraph = cellList.get(1).getParagraphArray(0);
        cellParagraphRun = cellParagraph.getRuns().get(0);
        cellParagraphRun.setText("数据类型");
        cellParagraph = cellList.get(2).getParagraphArray(0);
        cellParagraphRun = cellParagraph.getRuns().get(0);
        cellParagraphRun.setText("长度");
        cellParagraph = cellList.get(3).getParagraphArray(0);
        cellParagraphRun = cellParagraph.getRuns().get(0);
        cellParagraphRun.setText("小数位");
        cellParagraph = cellList.get(4).getParagraphArray(0);
        cellParagraphRun = cellParagraph.getRuns().get(0);
        cellParagraphRun.setText("默认值");
        cellParagraph = cellList.get(5).getParagraphArray(0);
        cellParagraphRun = cellParagraph.getRuns().get(0);
        cellParagraphRun.setText("说明");
//		cell = row.createCell(0);
//		cell.setCellStyle(style);
//		cell.setCellValue("列名");
//		cell = row.createCell(1);
//		cell.setCellStyle(style);
//		cell.setCellValue("数据类型");
//		cell = row.createCell(2);
//		cell.setCellStyle(style);
//		cell.setCellValue("长度");
//		cell = row.createCell(3);
//		cell.setCellStyle(style);
//		cell.setCellValue("小数位");
//		cell = row.createCell(4);
//		cell.setCellStyle(style);
//		cell.setCellValue("默认值");
//		cell = row.createCell(5);
//		cell.setCellStyle(style);
//		cell.setCellValue("说明");
        for(int i = 0; i < tableData.size(); i++) {
			TableInfo ti = tableData.get(i);
            cellList = rowList.get(i+2).getTableCells();
            cellParagraph = cellList.get(0).getParagraphArray(0);
            cellParagraphRun = cellParagraph.getRuns().get(0);
            cellParagraphRun.setText(String.valueOf(ti.getColumnName()));
            cellParagraph = cellList.get(1).getParagraphArray(0);
            cellParagraphRun = cellParagraph.getRuns().get(0);
            cellParagraphRun.setText(String.valueOf(ti.getDataType()));
            cellParagraph = cellList.get(2).getParagraphArray(0);
            cellParagraphRun = cellParagraph.getRuns().get(0);
            cellParagraphRun.setText(String.valueOf(ti.getDataLength()));
            cellParagraph = cellList.get(3).getParagraphArray(0);
            cellParagraphRun = cellParagraph.getRuns().get(0);
            cellParagraphRun.setText(String.valueOf(ti.getDataScale()));
            cellParagraph = cellList.get(4).getParagraphArray(0);
            cellParagraphRun = cellParagraph.getRuns().get(0);
            cellParagraphRun.setText(String.valueOf(ti.getDataDefault()));
            cellParagraph = cellList.get(5).getParagraphArray(0);
            cellParagraphRun = cellParagraph.getRuns().get(0);
            cellParagraphRun.setText(String.valueOf(ti.getcComments()));
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