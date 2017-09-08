import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;

/**
 * Created by dimanandhar on 9/8/17.
 */
public class MyProgressBar extends JFrame {
    JProgressBar progressBar;
    JButton stop;

    public MyProgressBar(){
        initFrame();
    }
    public void initFrame(){
        setTitle("Creating Excel File...");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setBorder(BorderFactory.createTitledBorder("Creating Excel.."));
        getContentPane().add(progressBar, BorderLayout.NORTH);
        stop = new JButton("Stop");
        getContentPane().add(stop, BorderLayout.SOUTH);

        setSize(300, 100);
        setVisible(true);
    }
}
