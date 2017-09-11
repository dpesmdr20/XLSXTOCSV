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
    String filePath = "/home/dimanandhar/dipesh/sample1.xlsx";
    double no_of_rows =1000000;
    private int no_of_col=1000;

    Cell cell;
    boolean stopped = false;
    MyProgressBar myProgressBar;
    JProgressBar progressBar;
    JButton stop;
    Border border;
    public CreateXlsFile(){
        myProgressBar = new MyProgressBar("Creating Excel File...");
        progressBar = myProgressBar.getProgressBar();
        stop = myProgressBar.stop;
        PerfomanceTracker.init();
        String dummyString ="i @M DuMmY...";
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
                    myProgressBar.stop.setEnabled(false);
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

