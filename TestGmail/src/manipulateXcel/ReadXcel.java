package manipulateXcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
//import org.hamcrest.core.IsNot;

public class ReadXcel  {
	
//	static File file = new File("C:\\Workspace\\TestGmail\\keyword.xlsx");
//	private static XSSFWorkbook wb = new XSSFWorkbook(fiStream);
//public Object getKeywordData(){

	public static void main(String[] args) throws IOException{
			//File file = new File("C:\\Workspace\\TestGmail\\keyword.xlsx");
			FileInputStream fis = new FileInputStream(new File("C:\\Workspace\\TestGmail\\keword.xls"));
			HSSFWorkbook wb = new HSSFWorkbook(fis);
			
		//Object objData[][] = new Object[5][5];
		Sheet sheet  = wb.getSheetAt(0);
		String strData[][] = new String[3][4]; 
		for (int i =1; i< sheet.getLastRowNum();i++){
			Row row = sheet.getRow(i);
			for (int j=0;j<2;j++){						                                       ;
			if((row.getCell(2).toString()).equalsIgnoreCase("Y")){
				//System.out.println("RunMode is \"Y\"");
				Sheet keywordSheet = wb.getSheet(row.getCell(0).toString());
				//System.out.println("**"+row.getCell(0).toString()+"&&" +keywordSheet.getLastRowNum()+"@@"+ keywordSheet.getSheetName());
				for (int k = 1;k<= keywordSheet.getLastRowNum();k++){
					//System.out.println("a");
					Row keywordRow = keywordSheet.getRow(k);					
					for (int l=0;l< keywordRow.getLastCellNum();l++){
						//System.out.println("b");
						//strData = new String[keywordSheet.getLastRowNum()][keywordRow.getLastCellNum()];
						k--;
						//System.out.println("b1");
						//System.out.println("&&" + keywordRow.getCell(l)+""+ k +":"+l);
						if (keywordRow.getCell(l) != null)
						strData[k][l] =  keywordRow.getCell(l).toString();k++;
						//System.out.println("b2");
						System.out.println("b2" +strData[k][l]);
					}
				}			
			}//else System.out.println("Run out of scene ");
			}			
		}
//		for (int i =0;i<4;i++){
//			System.out.println("C");
//			for (int j =0; j<9;j++){
//				System.out.println(strData[i][j]);
//			}
//		}
		//return objData;
	}	
}