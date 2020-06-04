package org.sluggard.oot;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.sluggard.oot.bean.TableInfo;
import org.sluggard.oot.dao.SimpleDao;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class OotApplicationTests {
	
	@Autowired
	private SimpleDao simpleDao;

	@Test
	void contextLoads() throws Exception {
		List<TableInfo> list = simpleDao.runSimpleSql();
		System.out.println(list.size());
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFCellStyle style = workbook.createCellStyle();
		style.setBorderBottom(BorderStyle.THIN); //�±߿�
		style.setBorderLeft(BorderStyle.THIN);//��߿�
		style.setBorderTop(BorderStyle.THIN);//�ϱ߿�
		style.setBorderRight(BorderStyle.THIN);//�ұ߿�
		Map<String, List<TableInfo>> map = new HashMap<>();
		list.forEach(t -> {
			System.out.println(t);
			if(!map.containsKey(t.getTableName())) {
				List<TableInfo> table = new ArrayList<>();
				map.put(t.getTableName(), table);
			}
			map.get(t.getTableName()).add(t);
		});
		map.forEach((k, v)->{
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
//			row.createCell(3).setCellValue(v.get(0).gettComments());
			// �ϲ�����ռ����(4���������ֱ�Ϊ��ʼ�У������У���ʼ�У�������)
	        // �к��ж��Ǵ�0��ʼ����������ʼ��������ϲ�
	        // �����Ǻϲ�excel�����ڵ�����Ϊһ��
	        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
	        sheet.addMergedRegion(region);
	        region = new CellRangeAddress(0, 0, 3, 5);
	        sheet.addMergedRegion(region);
//	        XSSFCell cell;
			row = sheet.createRow(1);
			cell = row.createCell(0);
			cell.setCellStyle(style);
			cell.setCellValue("����");
			cell = row.createCell(1);
			cell.setCellStyle(style);
			cell.setCellValue("��������");
			cell = row.createCell(2);
			cell.setCellStyle(style);
			cell.setCellValue("����");
			cell = row.createCell(3);
			cell.setCellStyle(style);
			cell.setCellValue("С��λ");
			cell = row.createCell(4);
			cell.setCellStyle(style);
			cell.setCellValue("Ĭ��ֵ");
			cell = row.createCell(5);
			cell.setCellStyle(style);
			cell.setCellValue("˵��");
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
			sheet.autoSizeColumn(0); //������һ�п��
	        sheet.autoSizeColumn(1); //�����ڶ��п��
	        sheet.autoSizeColumn(2); //���������п��
	        sheet.autoSizeColumn(3); //���������п��
	        sheet.autoSizeColumn(4); //���������п��
	        sheet.autoSizeColumn(5); //���������п��
//			setSizeColumn(sheet, 6);
		});
		File file = new File("aa.xlsx");
		FileOutputStream stream = new FileOutputStream(file);
        // ��Ҫ���쳣
        workbook.write(stream);
         //����
        stream.close();
	}
	
	private void setSizeColumn(XSSFSheet sheet, int size) {
        for (int columnNum = 0; columnNum < size; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                XSSFRow currentRow;
                //��ǰ��δ��ʹ�ù�
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
 
                if (currentRow.getCell(columnNum) != null) {
                    XSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == CellType.STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth);
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