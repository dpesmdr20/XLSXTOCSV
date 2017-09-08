import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.Random;

/**
 * Created by dimanandhar on 9/7/17.
 */
public class CreateXlsFile{
    String filePath = "/home/dimanandhar/dipesh/test.xlsx";
    double no_of_rows =1048575;
    Cell cell;
    Random random;
    boolean stopped = false;
    private int no_of_col=16383;
    MyProgressBar myProgressBar;
    JProgressBar progressBar;
    JButton stop;
    Border border;
    public CreateXlsFile(){
        myProgressBar = new MyProgressBar();
        progressBar = myProgressBar.progressBar;
        stop = myProgressBar.stop;
        PerfomanceTracker.init();
        random = new Random();
        String dummyString ="One morning";
        try {
            writeExcel(dummyString,filePath);
            PerfomanceTracker.stop();
            System.out.println("ExcelToCSV Memory Used: "+PerfomanceTracker.getMemoryUsage()+"  RunTime Period:"+PerfomanceTracker.getRunTimePeriod());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public void writeExcel(String dummyString,String excelFilePath) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
       // BufferedOutputStream bufferedOutputStream=new BufferedOutputStream(outputStream);
        //byte[] buff = new byte[32 * 1024];

        SXSSFWorkbook workbook = new SXSSFWorkbook();
        Sheet shits = null;

        stop.addActionListener(new ActionListener()
        {
            public void actionPerformed(ActionEvent e)
            {
               stopped = true;
            }
        });
            Row row;
            shits = workbook.createSheet();
            for(int i=0;i<no_of_rows;i++){
                if(!stopped) {
                    row = shits.createRow(i);
                    progressBar.setValue((int) getProgress(i));
                    writeBook(dummyString, row);
                }
                else
                    break;
            }

        border = BorderFactory.createTitledBorder("Writing Excel..");
        progressBar.setBorder(border);

        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        progressBar.setValue(100);
        myProgressBar.dispose();
        long size = new FileInputStream(filePath).getChannel().size();
        JOptionPane.showMessageDialog(null,"Sucessfully created "+PerfomanceTracker.bytesToMegabytes(size) +" mb Excel File.");
    }
    private void writeBook(String dummyString, Row row) {
        for(int i=0;i<no_of_col;i++){
            cell = row.createCell(i);
            cell.setCellValue(dummyString);
        }
    }
    public double getProgress(int cur){
        return cur*100/no_of_rows;
    }
}

