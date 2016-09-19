import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;


public class Main {

    public static void main(String[] args) {

        try
        {
            FileInputStream file = new FileInputStream(new File("C:\\CodeCool\\numbers.xlsx"));

            //Create Workbook instance holding reference to .xlsx file
            Workbook workbook = new XSSFWorkbook (file);

            //Get first/desired sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            StringBuilder eredmeny = new StringBuilder();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            //System.out.print(cell.getNumericCellValue() );
                            eredmeny.append((int) cell.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_STRING:
                            //System.out.print(cell.getStringCellValue() );
                            break;
                    }
                }

            }
            file.close();
            System.out.println(eredmeny);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }



    }
}
