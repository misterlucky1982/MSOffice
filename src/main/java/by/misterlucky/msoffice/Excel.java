package by.misterlucky.msoffice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Excel {
	
	private static final String ORIGIN = "sources/origin.xls";

	public static boolean write(String fileName, String[][] lines) throws ExcelException{
		File file = new File(fileName);
		if(!file.exists())file=createCopyOfFile(fileName);
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(file);
		} catch (FileNotFoundException e1) {
			throw new ExcelException();
		}
		HSSFWorkbook myWorkBook = null;
		try {
			myWorkBook = new HSSFWorkbook(fis);
		} catch (IOException e1) {
			throw new ExcelException();
		}
		HSSFSheet mySheet = myWorkBook.getSheetAt(0);
		Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
		for (int index = 0; index < lines.length; index++) {
			data.put(index, lines[index]);
		}
		int rownum = mySheet.getLastRowNum();
		for (int key = 0; key < lines.length; key++) {
			Row row = mySheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Boolean) {
					cell.setCellValue((Boolean) obj);
				} else if (obj instanceof Date) {
					cell.setCellValue((Date) obj);
				} else if (obj instanceof Double) {
					cell.setCellValue((Double) obj);
				}
			}
		}
		FileOutputStream os = null;
		try {
			os = new FileOutputStream(file);
		} catch (FileNotFoundException e) {
			throw new ExcelException();
		}
		try {
			myWorkBook.write(os);
		} catch (IOException e) {
			throw new ExcelException();
		}
		return true;
}
	
	public static boolean write(File destination, String[][] lines) throws ExcelException{
		if(destination==null||!destination.exists())throw new ExcelException("invalid destination file");
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(destination);
		} catch (FileNotFoundException e1) {
			return false;
		}
		HSSFWorkbook myWorkBook = null;
		try {
			myWorkBook = new HSSFWorkbook(fis);
		} catch (IOException e1) {
			throw new ExcelException();
		}
		HSSFSheet mySheet = myWorkBook.getSheetAt(0);
		Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
		for (int index = 0; index < lines.length; index++) {
			data.put(index, lines[index]);
		}
		int rownum = mySheet.getLastRowNum();
		for (int key = 0; key < lines.length; key++) {
			Row row = mySheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Boolean) {
					cell.setCellValue((Boolean) obj);
				} else if (obj instanceof Date) {
					cell.setCellValue((Date) obj);
				} else if (obj instanceof Double) {
					cell.setCellValue((Double) obj);
				}
			}
		}
		FileOutputStream os = null;
		try {
			os = new FileOutputStream(destination);
		} catch (FileNotFoundException e) {
			throw new ExcelException();
		}
		try {
			myWorkBook.write(os);
		} catch (IOException e) {
			throw new ExcelException();
		}
		return true;
}
	
	public static String readString(String fileName, int sheet, int rOw, int ceLL){
		return Excel.readString(new File(fileName), sheet, rOw, ceLL);
	}
	
	public static String readString(File source, int sheet, int rOw, int ceLL) {
		String result = null;
		try{
			HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(source));
			HSSFSheet myExcelSheet = myExcelBook.getSheetAt(sheet);
			if(rOw>myExcelSheet.getLastRowNum())return null;
			Iterator<org.apache.poi.ss.usermodel.Row> rowIterator = myExcelSheet.iterator();
			while(rowIterator.hasNext()){
				org.apache.poi.ss.usermodel.Row row = (org.apache.poi.ss.usermodel.Row)
						rowIterator.next();		
				if(((org.apache.poi.ss.usermodel.Row) row).getRowNum()==rOw){
					
				Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = 
						((org.apache.poi.ss.usermodel.Row) row).cellIterator();
				while(cellIterator.hasNext()){   
					org.apache.poi.ss.usermodel.Cell cell = 
							(org.apache.poi.ss.usermodel.Cell) cellIterator.next();
					if(((org.apache.poi.ss.usermodel.Cell) cell).getColumnIndex()==ceLL){
						if(((org.apache.poi.ss.usermodel.Cell) cell).getCellType()
								==HSSFCell.CELL_TYPE_STRING){
							result = ((org.apache.poi.ss.usermodel.Cell) cell)
									.getStringCellValue();
							return result;
							}else{
								if(((org.apache.poi.ss.usermodel.Cell) cell)
										.getCellType()==HSSFCell.CELL_TYPE_NUMERIC){
									result = ""+((org.apache.poi.ss.usermodel.Cell) cell)
											.getNumericCellValue();
									return result;
									}else{
										if(((org.apache.poi.ss.usermodel.Cell) cell)
												.getCellType()==HSSFCell.CELL_TYPE_FORMULA){
											try{										
												result = ""+((org.apache.poi.ss.usermodel.Cell) cell)
														.getNumericCellValue();
											}catch(java.lang.IllegalStateException exc){
												return null;
											}																					
											
											return result;
										}else return null;
									}
						}
					}
				if(((org.apache.poi.ss.usermodel.Cell) cell).getColumnIndex()>ceLL)return null;
			}
				}
			if(((org.apache.poi.ss.usermodel.Row) row).getRowNum()>rOw)return null;
			}
		}catch(FileNotFoundException e){
			return null;
		}catch(IOException ee){
			return null;
		}
		return result;
}
	
	
	private static File createCopyOfFile(String fileName) {
		try{
			File dest = new File(fileName);
			if(!dest.exists())dest.createNewFile();
			InputStream is = null;
		    OutputStream os = null;
		    try {
		        is = new FileInputStream(ORIGIN);
		        os = new FileOutputStream(dest);
		        byte[] buffer = new byte[1024];
		        int length;
		        while ((length = is.read(buffer)) > 0) {
		            os.write(buffer, 0, length);
		        }
		    }finally {
		        if(is!=null)is.close();
		        if(os!=null)os.close();
		    }
		}catch(IOException e){
			return null;
		}
		return new File(fileName);
}
}
