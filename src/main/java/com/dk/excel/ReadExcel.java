package com.dk.excel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	@SuppressWarnings("resource")
	public static void main(String[] args) {
		try {
			List<List<String>> infos = new ArrayList<List<String>>();
			InputStream is = new FileInputStream("D:/data.xlsx");
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
			XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
			for(int i=1;i<=xssfSheet.getLastRowNum();i++){
				XSSFRow xssfRow = xssfSheet.getRow(i);
				int minCell = xssfRow.getFirstCellNum();
				int maxCell = xssfRow.getLastCellNum();
				List<String> rowList = new ArrayList<String>();
				for(int j=minCell;j<maxCell;j++){
					XSSFCell cell = xssfRow.getCell(j);
					rowList.add(cell.toString());
				}
				infos.add(rowList);
			}
			for(List<String> lists:infos){
				for(String str:lists){
					System.out.println(str);
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
