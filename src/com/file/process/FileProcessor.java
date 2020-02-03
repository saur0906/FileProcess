package com.file.process;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileProcessor {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		String input = "E:\\projects\\p_proj\\source\\input.csv";
		String template = "E://projects//p_proj//template.xlsx";
		String result = "E:\\projects\\p_proj\\target\\result.xlsx";
		
		File template_file = new File(template);
		File res_file = new File(result);
		//create a  copy of template in target directory
		Files.copy(template_file.toPath(),new File(result).toPath(),StandardCopyOption.REPLACE_EXISTING);
		
		FileInputStream fis = new FileInputStream(res_file);
		
        //Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        //Get first/desired sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);
		
        //Read input file line by line
        BufferedReader br = null;
        String line = "";
        String cvsSplitBy = ",";
        
        br = new BufferedReader(new FileReader(input));
        int line_no =1;
        while ((line = br.readLine()) != null) {

            // use comma as separator
            String[] lineOfInputCsv = line.split(cvsSplitBy);
            
            if (line_no == 1) {
				continue;//skip processing the first line
			}
            
            line_no++;
            
            String textToBeFound = "("+lineOfInputCsv[0]+")";
            
            //start search of index in result.xlsx , for example (1)
            
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            int rowcount = 1;
            while (rowIterator.hasNext()) {
            	
            	if(rowcount>7){
            		continue;//we only want to check through first 7 rows of the sheet
            	}
            	
				Row row = (Row) rowIterator.next();
				
				Cell cell = row.getCell(0);//gets the first cell of each row which contains our regex we want to find
				
				String cellValue = cell.getStringCellValue();
				
				if (cellValue.contains(textToBeFound)){//check if the row has something like (1) or (2)
					
					// now iterate through the content of lineOfInputCsv from index 1 and place them in subsequent cells
					for( int i=1;i<lineOfInputCsv.length;i++){
						Cell cell_to_be_modified = row.getCell(i);
						cell_to_be_modified.setCellValue(Integer.parseInt(lineOfInputCsv[i]));
					}//for loop close
					
				}
				
				rowcount++;
			}//row iterator while ends


        }//while loop through lines of csv ends
		
        fis.close();
        br.close();
        
	}

}
