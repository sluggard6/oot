package org.sluggard.oot;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellRangeAddress;
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
			row.createCell(0).setCellValue(k);
			row.createCell(3).setCellValue(v.get(0).gettComments());
			// �ϲ�����ռ����(4���������ֱ�Ϊ��ʼ�У������У���ʼ�У�������)
	        // �к��ж��Ǵ�0��ʼ����������ʼ��������ϲ�
	        // �����Ǻϲ�excel�����ڵ�����Ϊһ��
	        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
	        sheet.addMergedRegion(region);
	        region = new CellRangeAddress(0, 0, 3, 5);
	        sheet.addMergedRegion(region);
			
			row = sheet.createRow(1);
			row.createCell(0).setCellValue("����");
			row.createCell(1).setCellValue("��������");
			row.createCell(2).setCellValue("����");
			row.createCell(3).setCellValue("С��λ");
			row.createCell(4).setCellValue("Ĭ��ֵ");
			row.createCell(5).setCellValue("˵��");
			for(int i=0;i<v.size();i++) {
				row = sheet.createRow(i+2);
				TableInfo ti = v.get(i);
				row.createCell(0).setCellValue(ti.getColumnName());
				row.createCell(1).setCellValue(ti.getDataType());
				row.createCell(2).setCellValue(ti.getDataLength());
				row.createCell(3).setCellValue(ti.getDataScale());
				row.createCell(4).setCellValue(ti.getDataDefault());
				row.createCell(5).setCellValue(ti.getcComments());
			}
		});
		File file = new File("aa.xlsx");
		FileOutputStream stream = new FileOutputStream(file);
        // ��Ҫ���쳣
        workbook.write(stream);
         //����
        stream.close();
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