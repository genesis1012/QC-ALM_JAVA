package HP;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelApiTest {
	
	public FileInputStream fis = null;
    public HSSFWorkbook workbook = null;
    public HSSFSheet sheet = null;
    public HSSFRow row = null;
    public HSSFCell cell = null;
    public String sheetName = "BUG";
    public Boolean close=null;
    
    /******************Super contructor para leer y escribir documento de Excel*/
    public ExcelApiTest(String xlFilePath) throws Exception
    {
        fis = new FileInputStream(xlFilePath);
        workbook = new HSSFWorkbook(fis);
        fis.close();
    }
    
    /********************Validar si el archivo de defectos esta abierto****************/
    public void Excel_read(String yourFile) throws IOException{
        File file = new File(yourFile);
        File sameFileName = new File(yourFile);
        if(file.renameTo(sameFileName)){
            System.out.println("Archivo de defectos esta cerrado");
            close=true;
        }else{
        	System.out.println("Archivo de defectos esta abierto");
            close=false;
        }
    }
    
    /*****************Buscar la fila en la que se encuentra el id del defecto
     * @throws IOException *************/
    public int find_row(String colName, String Val_id) throws IOException{
    	int cel_Num = -1;
	    	try
	        {
	            int col_Num = -1;
	            sheet = workbook.getSheet(sheetName);// getSheet(sheetName);
	            row = sheet.getRow(0);
	            for(int i = 0; i < row.getLastCellNum(); i++)
	            {
	                if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
	                    col_Num = i;
	            }
	            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
	            	  row = sheet.getRow(rowIndex);
	            	  if (row != null) {
	            	    Cell cell = row.getCell(col_Num);
	            	    if (cell != null) {
	            	    	if(Campo(cell).equals(Val_id)){
	            	    		cel_Num = rowIndex;
	            	    	}
	            	    	
	            	    }
	            	  }
	            	}
	            
	        }
	        catch(Exception e)
	        {
	            e.printStackTrace();
	            //return "row  or column "+colName +" does not exist  in Excel";
	        }
    	return cel_Num;
    }
    
    /********************Devuelve el valor del campo identificando la fila del defecto************/
    public String getCellData(String colName, String Val_id)
    {
        try
        {
            int col_Num = -1;
            int cel_Num = find_row("Val_ID", Val_id);
            sheet = workbook.getSheet(sheetName);// getSheet(sheetName);
            row = sheet.getRow(0);
            for(int i = 0; i < row.getLastCellNum(); i++)
            {
                if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
                    col_Num = i;
            }
            
            row = sheet.getRow(cel_Num);
            cell = row.getCell(col_Num);
            return Campo(cell);
        }
        catch(Exception e)
        {
            e.printStackTrace();
            return "row  or column "+colName +" does not exist  in Excel";
        }
    }
    
    /******************Escribe en el id del bug en la fila del fefecto********************/
    public void writeXLSXFile(String path, String Val_id, String colName, String bug_id) throws IOException {
        try {
        	int col_Num = -1;
            int cel_Num = find_row("Val_ID", Val_id);
            sheet = workbook.getSheet(sheetName);
            row = sheet.getRow(0);
            for(int i = 0; i < row.getLastCellNum(); i++)
            {
                if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
                    col_Num = i;
            }
            row = sheet.getRow(cel_Num);
            cell = row.getCell(col_Num);
            if (cell != null)
                cell = row.createCell(col_Num);
            cell.setCellValue(bug_id);

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream(path);
            workbook.write(fileOut);
            fileOut.close();

        } catch (FileNotFoundException e) {
        	e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (EncryptedDocumentException e) {
			e.printStackTrace();
		}
    }
    
    /************Identifica de que tipo es el campo que se va a devolver****************/
    private String Campo(Cell cell){
    	if(cell.getCellType() == Cell.CELL_TYPE_STRING){
            return cell.getStringCellValue();
    	}
        else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_FORMULA)
        {
        	int i = (int)cell.getNumericCellValue();
            String cellValue = String.valueOf(i);
            if(HSSFDateUtil.isCellDateFormatted(cell))
            {
                DateFormat df = new SimpleDateFormat("dd/MM/yy");
                Date date = cell.getDateCellValue();
                cellValue = df.format(date);
            }
            return cellValue;
        }else if(cell.getCellType() == Cell.CELL_TYPE_BLANK){
            return "";
        }
        else{
            return String.valueOf(cell.getBooleanCellValue());
        }
    }

}
