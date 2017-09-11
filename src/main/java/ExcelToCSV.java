
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * Created by dimanandhar on 9/6/17.
 */
public class ExcelToCSV {
     String PATH = "/home/dimanandhar/dipesh/sample1.xlsx";
    public ExcelToCSV(){
        InputStream is = null;
        Workbook workbook = null;
        try {
            is = new FileInputStream(new File(PATH));
            workbook = StreamingReader.builder()
                    .rowCacheSize(10)    // number of rows to keep in memory (defaults to 10)
                    .bufferSize(6144)     // buffer size to use when reading InputStream to file (defaults to 1024)
                    .open(is);            // InputStream or File for XLSX file (required)
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Sheet sheet = workbook.getSheetAt(0);
        writeToCSV(sheet);
    }

    public void writeToCSV(Sheet sheet) {
        BufferedWriter bw = null;
        FileWriter fw = null;
        try {
            File file = new File("/home/dimanandhar/dipesh/sample2.csv");
            // if file doesnt exists, then create it
            if (!file.exists()) {
                file.createNewFile();
            }

            // true = append file
            fw = new FileWriter(file.getAbsoluteFile(), true);
            bw = new BufferedWriter(fw);
            Row x = sheet.getRow(10);
          System.out.println(x.getCell(1).getStringCellValue());
        for (Row r : sheet) {
            PerfomanceTracker.init();
            for (Cell c : r) {
                    bw.write("\"" + c.getStringCellValue() + "\",");
                }
            PerfomanceTracker.stop();
            System.out.println("ExcelToCSV Memory Used: "+PerfomanceTracker.getMemoryUsage()+"  RunTime Period:"+PerfomanceTracker.getRunTimePeriod());
            }

        } catch (IOException e) {
            e.printStackTrace();

        } finally {

            try {

                if (bw != null)
                    bw.close();

                if (fw != null)
                    fw.close();

            } catch (IOException ex) {

                ex.printStackTrace();

            }
        }
    }

}
