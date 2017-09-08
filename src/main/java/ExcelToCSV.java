
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created by dimanandhar on 9/6/17.
 */
public class ExcelToCSV {
    static StringBuffer csv = null;
    public static void echoAsCSV(Sheet sheet) {
        csv = new StringBuffer();
        Row row = null;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                csv.append("\"" + row.getCell(j) + "\",");
            }
            csv.append("\n");
        }
    }
    public ExcelToCSV(){
        PerfomanceTracker.init();

        InputStream inp = null;
        try {
            inp = new FileInputStream("/home/dimanandhar/dipesh/test.xlsx");
            Workbook wb = WorkbookFactory.create(inp);

            for(int i=0;i<wb.getNumberOfSheets();i++) {
                System.out.println(wb.getSheetAt(i).getSheetName());
                echoAsCSV(wb.getSheetAt(i));
            }
            PerfomanceTracker.stop();
            System.out.println("ExcelToCSV Memory Used: "+PerfomanceTracker.getMemoryUsage()+"  RunTime Period:"+PerfomanceTracker.getRunTimePeriod());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ExcelToCSV.class.getName()).log(Level.SEVERE, null, ex);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelToCSV.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelToCSV.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                inp.close();
            } catch (IOException ex) {
                Logger.getLogger(ExcelToCSV.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }


}
