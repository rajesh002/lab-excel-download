package service;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


import model.Prograd;
public class ExcelGenerator {
	
	FileOutputStream out;
	public HSSFWorkbook excelGenerate(Prograd prograd, List<Prograd> list) throws IOException {
		try {
			HSSFWorkbook hwb=new HSSFWorkbook();
			HSSFSheet sheet = hwb.createSheet("Prograd List");
			int rowNum = 1;
			for(int index=0;index<list.size();index++) {
				HSSFRow row = sheet.createRow(rowNum++);

	            row.createCell(0).setCellValue(prograd.getName());

	            row.createCell(1).setCellValue(prograd.getId());

	            row.createCell(2).setCellValue(prograd.getRate());
	            
	            row.createCell(3).setCellValue(prograd.getRecommend());
	            
	            row.createCell(4).setCellValue(prograd.getComment());
			}
			// Do not modify the lines given below
			 out = new FileOutputStream("E:\\proGrad\\Week3\\Day6\\lab-excel-download\\src\\progradData.xlsx");
			hwb.write(out);
			return hwb;
			}
		catch (Exception e) {
				e.printStackTrace();
			}
		finally {
			out.close();
		}
		return null;
		
	}
}
