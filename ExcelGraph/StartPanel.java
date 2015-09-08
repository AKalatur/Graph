package ExcelGraph;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class StartPanel extends JPanel{
    JLabel TOBOLabel;
    public StartPanel() throws IOException {

        JLabel label = new JLabel("TOBO");
        add(label);

        // Выпадающий список с перечнем кодов отделений
        Object[] items = getListOfTOBO();
        final JComboBox comboBox = new JComboBox(items);
        Graph.TOBO = Integer.valueOf((String)comboBox.getSelectedItem());
        add(comboBox);

        comboBox.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    Graph.TOBO = Integer.valueOf((String)comboBox.getSelectedItem());
                    TOBOLabel.setText(getTOBOName());
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        });

        //кнопка для запуска построения графика
        JButton okButton = new JButton("OK");
        add(okButton);
        okButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                Graph.TOBO = Integer.valueOf((String)comboBox.getSelectedItem());
                Graph.TOBOName = TOBOLabel.getText();
                Thread t = new Thread(new Graph());
                t.start();
            }
        });

        //наименование отделения
        TOBOLabel = new JLabel(getTOBOName());
        TOBOLabel.setPreferredSize(new Dimension(190, 50));
        TOBOLabel.setHorizontalAlignment(SwingConstants.CENTER);
        add(TOBOLabel);
    }

    //получение списка номеров отделений из файла
    public static Object[] getListOfTOBO() throws IOException {
        ArrayList<String> sList = new ArrayList<String>();
        FileInputStream fis = new FileInputStream(Graph.fileName);
        Workbook wb = new HSSFWorkbook(fis);

        for (Row row : wb.getSheetAt(0)) {
            sList.add(row.getCell(0).getStringCellValue());
        }

        sList.remove(0);
        fis.close();

        return sList.toArray();
    }

        //определение наименования выбраного отделения в выпадающем списке
    public static String getTOBOName() throws IOException {
        FileInputStream fis = new FileInputStream(Graph.fileName);
        Workbook wb = new HSSFWorkbook(fis);

        for (Row row : wb.getSheetAt(0)) {
            if (row.getCell(0).getStringCellValue().equals(String.valueOf(Graph.TOBO)))
            return row.getCell(1).getStringCellValue();
        }
        return null;
    }
}