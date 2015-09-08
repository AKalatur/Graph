package ExcelGraph;

import javax.swing.*;
import java.io.IOException;

class StartFrame extends JFrame {

    public StartFrame() throws IOException {
        setBounds(300, 300, 300, 150);
        setTitle("START WINDOW");
        setResizable(false);

        StartPanel sPanel = new StartPanel();
        add(sPanel);
    }
}