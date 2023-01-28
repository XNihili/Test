package test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

public class ExcelWriter {

	private XSSFWorkbook myWorkBookXSSF = null;
	private XSSFSheet mySheetXSSF = null;
	private HSSFWorkbook myWorkBookHSSF = null;
	private HSSFSheet mySheetHSSF = null;
	
	/**
	 * private boolean checkInExcelFileXSSF(String sujet)
	 * Vérifie si le sujet est présent dans la 1ere colonne de la feuille du fichier excel
	 * retourne true s'il est présent, false sinon
	 */
	private boolean checkInExcelFileXSSF(String sujet) {
		try {
	        Iterator<Row> rowIterator = this.mySheetXSSF.iterator();
	        while (rowIterator.hasNext()) {
	            Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                if(cellIterator.hasNext() ) {
                	cellIterator.next();
	                Cell cell = cellIterator.next();
	                if(sujet.equals(cell.getStringCellValue()))
	                	return true;
                }
	        }
		
		} catch (Exception e) {
			System.out.println("Exception in checking data in Excel tab : "+e.getMessage());
		}
		return false;
	}

	/**
	 * private boolean checkInExcelFileHSSF(String sujet)
	 * Vérifie si le sujet est présent dans la 1ere colonne de la feuille du fichier excel
	 * retourne true s'il est présent, false sinon
	 */
	private boolean checkInExcelFileHSSF(String sujet) {
		try {
	        Iterator<Row> rowIterator = this.mySheetHSSF.iterator();
	        while (rowIterator.hasNext()) {
	            Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                if(cellIterator.hasNext() ) {
                	cellIterator.next();
	                Cell cell = cellIterator.next();
	                if(sujet.equals(cell.getStringCellValue()))
	                	return true;
                }
	        }
		
		} catch (Exception e) {
			System.out.println("Exception in checking data in Excel tab : "+e.getMessage());
		}
		return false;
	}

	/**
	 * private void openExcelFileXSSF(String file)
	 * ouvre fichier excel dont le nom est le paramètre
	 */
	private void openExcelFileXSSF(String file) {
		try {
			FileInputStream fis = new FileInputStream(file);
			this.myWorkBookXSSF = new XSSFWorkbook(fis);
			this.mySheetXSSF = this.myWorkBookXSSF.getSheetAt(1);
		} catch (Exception e) {
			System.out.println("Exception in opening Excel file : "+e.getMessage());
		}
	}
	/**
	 * private void openExcelFileHSSF(String file)
	 * ouvre fichier excel dont le nom est le paramètre
	 */	
	private void openExcelFileHSSF(String file) {
		try {		
			FileInputStream fis = new FileInputStream(file);
			this.myWorkBookHSSF = new HSSFWorkbook(fis);
			this.mySheetHSSF = this.myWorkBookHSSF.getSheetAt(1);
		} catch (Exception e) {
			System.out.println("Exception in opening Excel file : "+e.getMessage());
		}
	}
	
	/**
	 * private void closeExcelFileXSSF(String file)
	 * sauvegarde les changements et ferme fichier excel dont le nom est le paramètre
	 */
	private void closeExcelFileXSSF(String file) {
		try {
			FileOutputStream fos = new FileOutputStream(file);
			this.myWorkBookXSSF.write(fos);
			this.myWorkBookXSSF.close();
			fos.close();
		} catch (Exception e) {
			System.out.println("Exception in closing Excel file : "+e.getMessage());
		}
	}
	
	/**
	 * private void closeExcelFileHSSF(String file)
	 * sauvegarde les changements et ferme fichier excel dont le nom est le paramètre
	 */	
	private void closeExcelFileHSSF(String file) {
		try {	
			FileOutputStream fos = new FileOutputStream(file);
			this.myWorkBookHSSF.write(fos);
			this.myWorkBookHSSF.close();
			fos.close();
		} catch (Exception e) {
			System.out.println("Exception in closing Excel file : "+e.getMessage());
		}
	}
	
	
	/**
	 * private void writeInExcelFileXSSF(String type, String sujet)
	 * Write sujet on the last row of the tab of the xls file
	 */
	private void writeInExcelFileXSSF(String type, String sujet) {
		int rownum = this.mySheetXSSF.getLastRowNum();
		rownum++;
        Map<String, Object[]> data = new HashMap<String, Object[]>();
        data.put(Integer.toString(rownum), new Object[] {type, sujet, "", "", "", ""});
		Set<String> newRows = data.keySet();
        for (String key : newRows) {
            // Creating a new Row in existing XLSX sheet
            Row row = this.mySheetXSSF.createRow(rownum);
            Object [] objArr = data.get(key);
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
	}

	/**
	 * private void writeInExcelFileHSSF(String type, String sujet)
	 * Write sujet on the last row of the tab of the xls file
	 */
	private void writeInExcelFileHSSF(String type, String sujet) {
		int rownum = this.mySheetHSSF.getLastRowNum();
        Map<String, Object[]> data = new HashMap<String, Object[]>();
        data.put("Mail", new Object[] {sujet, "", "", "", ""});
		Set<String> newRows = data.keySet();
        for (String key : newRows) {
            // Creating a new Row in existing XLSX sheet
            Row row = this.mySheetHSSF.createRow(rownum++);
            Object [] objArr = data.get(key);
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
	}
	
	/**
	 * public void insertDataInExcelFileXSSF(String file, String type, String sujet)
	 * look if the sujet is in the file, if not it is inserted at the last place on the 2nd tab 
	 */
	public void insertDataInExcelFileXSSF(String file, String type, String sujet) {
		this.openExcelFileXSSF(file);
		if(!this.checkInExcelFileXSSF(sujet)) {
			this.writeInExcelFileXSSF(type, sujet);
		}
		closeExcelFileXSSF(file);
	}

	/**
	 * public void insertDataInExcelFileHSSF(String file, String type, String sujet)
	 * look if the sujet is in the file, if not it is inserted at the last place on the 2nd tab 
	 */
	public void insertDataInExcelFileHSSF(String file, String type, String sujet) {
		this.openExcelFileHSSF(file);
		if(!this.checkInExcelFileHSSF(sujet)) {
			this.writeInExcelFileHSSF(type, sujet);
		}
		closeExcelFileHSSF(file);
	}
	
}
