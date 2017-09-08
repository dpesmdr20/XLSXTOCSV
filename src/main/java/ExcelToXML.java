import jxl.Sheet;
import jxl.Workbook;
import org.w3c.dom.Document;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.io.File;

/**
 * Created by dimanandhar on 9/6/17.
 */
public class ExcelToXML {
    File excelFile = new File("/home/dimanandhar/Downloads/SampleXLSFile_904kb.xls");
    JTable previewTable = null;

    public ExcelToXML() {
        PerfomanceTracker.init();
        // Create model for excel file
        if (excelFile.exists()) {
            try {
                Workbook workbook = Workbook.getWorkbook(excelFile);
                Sheet sheet = workbook.getSheets()[0];

                TableModel model = new DefaultTableModel(sheet.getRows(), sheet.getColumns());
                for (int row = 0; row < sheet.getRows(); row++) {
                    for (int column = 0; column < sheet.getColumns(); column++) {
                        String content = sheet.getCell(column, row).getContents();
                        model.setValueAt(content, row, column);
                    }
                }
                PerfomanceTracker.stop();
                System.out.println("ExcelToXML Memory Used: "+PerfomanceTracker.getMemoryUsage()+"  RunTime Period:"+PerfomanceTracker.getRunTimePeriod());

              //  previewTable.setModel(model);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "Error: " + e);
            }

        } else {
            JOptionPane.showMessageDialog(null, "File does not exist");
        }
    }
    public void xmlToCSV() throws Exception {
        File stylesheet = new File("/home/dimanandhar/Downloads/SampleXLSFile_904kb.xls");
        File xmlSource = new File("src/main/resources/data.xml");

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document document = builder.parse(xmlSource);

        StreamSource stylesource = new StreamSource(stylesheet);
        Transformer transformer = TransformerFactory.newInstance()
                .newTransformer(stylesource);
        Source source = new DOMSource(document);
        Result outputTarget = new StreamResult(new File("/tmp/x.csv"));
        transformer.transform(source, outputTarget);
    }
}
