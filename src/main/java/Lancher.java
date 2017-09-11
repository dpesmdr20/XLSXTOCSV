import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.xml.sax.SAXException;

import javax.swing.*;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.IOException;

/**
 * Created by dimanandhar on 9/6/17.
 */
public class Lancher {
    static String PATH = "/home/dimanandhar/dipesh/sample1.xlsx";
    public static void main(String[] args){
        new ExcelToCSV();
        // new CreateXlsFile();
        //new ExcelReader().read();
    }
}