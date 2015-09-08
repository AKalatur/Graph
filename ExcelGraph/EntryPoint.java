package ExcelGraph;

import javax.swing.*;
import java.io.IOException;

public class EntryPoint {
    public static void main(String[] args) throws IOException {
        //Graph.fileName = args[0];
        Graph.fileName = "SourceFile.xls";

        StartFrame sFrame = new StartFrame();
        sFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        sFrame.setVisible(true);

    }
}