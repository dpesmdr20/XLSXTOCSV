import javax.swing.*;
import java.awt.*;

/**
 * Created by dimanandhar on 9/8/17.
 */
public class MyProgressBar extends JFrame {
    private JProgressBar progressBar;
    JButton stop;
    public MyProgressBar(String title){
        initFrame(title);
    }
    public void initFrame(String title){
        setTitle(title);
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
    public JProgressBar getProgressBar() {
        return progressBar;
    }
}
