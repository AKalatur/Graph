package ExcelGraph;

import com.xeiam.xchart.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Locale;

public class Graph implements Runnable {
	static int TOBO;
	static String TOBOName;
	static String graphTitle;
	static String fileName;

	 public void run() {

		try {
			getTitle();

		 // создание графика
		 	Chart chart = new Chart(800, 600);

		// настройка визуализации графика

			Collection<Date> xData = getListOfX();
			Collection<Double> yData = getListOfY();

			chart.setChartTitle(TOBOName);
			chart.setXAxisTitle("Період/Дата");
			chart.setYAxisTitle("Сума");

			chart.getStyleManager().setPlotBackgroundColor(ChartColor.getAWTColor(ChartColor.GREY));
			chart.getStyleManager().setPlotGridLinesColor(new Color(255, 255, 255));
			chart.getStyleManager().setChartBackgroundColor(Color.WHITE);
			chart.getStyleManager().setLegendBackgroundColor(Color.PINK);
			chart.getStyleManager().setChartFontColor(Color.MAGENTA);
			chart.getStyleManager().setChartTitleBoxBackgroundColor(new Color(0, 222, 0));
			chart.getStyleManager().setChartTitleBoxVisible(true);
			chart.getStyleManager().setChartTitleBoxBorderColor(Color.BLACK);
			chart.getStyleManager().setPlotGridLinesVisible(false);
			chart.getStyleManager().setAxisTickPadding(20);
			chart.getStyleManager().setAxisTickMarkLength(15);
			chart.getStyleManager().setPlotPadding(20);
			chart.getStyleManager().setChartTitleFont(new Font(Font.MONOSPACED, Font.BOLD, 24));
			chart.getStyleManager().setLegendFont(new Font(Font.SERIF, Font.PLAIN, 18));
			chart.getStyleManager().setLegendPosition(StyleManager.LegendPosition.OutsideE);
			chart.getStyleManager().setLegendSeriesLineLength(5);
			chart.getStyleManager().setAxisTitleFont(new Font(Font.SANS_SERIF, Font.ITALIC, 18));
			chart.getStyleManager().setAxisTickLabelsFont(new Font(Font.SERIF, Font.PLAIN, 11));
			chart.getStyleManager().setDatePattern("dd/MM");
			chart.getStyleManager().setXAxisTickMarkSpacingHint(30);
			//chart.getStyleManager().setXAxisTicksVisible(false);

			chart.getStyleManager().setLocale(Locale.GERMAN);

			Series series = chart.addSeries(graphTitle, xData, yData);
			series.setLineColor(SeriesColor.BLUE);
			series.setMarkerColor(Color.ORANGE);
			series.setMarker(SeriesMarker.CIRCLE);
			series.setLineStyle(SeriesLineStyle.SOLID);

			//отображение графика
			new SwingWrapper(chart).displayChart();
		}
		catch (IOException e) {
			 e.printStackTrace();
		 }
	 }

	//получение листа для оси ординат
	public static ArrayList<Double> getListOfY() throws IOException {
		ArrayList<Double> dList = new ArrayList<Double>();

		FileInputStream fis = new FileInputStream(fileName);
		Workbook wb = new HSSFWorkbook(fis);

		for (int i = 1; i < wb.getSheetAt(0).getPhysicalNumberOfRows(); i++) {
			for (int j = 2; j < wb.getSheetAt(0).getRow(i).getPhysicalNumberOfCells(); j++) {
				if (TOBO == Integer.parseInt(wb.getSheetAt(0).getRow(i).getCell(0).getStringCellValue())) {
					dList.add(wb.getSheetAt(0).getRow(i).getCell(j).getNumericCellValue());
				}
			}
		}

		fis.close();

		return dList;
	}

	//получение листа для оси абсцисс
	public static ArrayList<Date> getListOfX() throws IOException {
		ArrayList<Date> dateList = new ArrayList<Date>();

		FileInputStream fis = new FileInputStream(fileName);
		Workbook wb = new HSSFWorkbook(fis);

		for (int i = 2; i < wb.getSheetAt(0).getRow(0).getPhysicalNumberOfCells(); i++) {
			dateList.add(wb.getSheetAt(0).getRow(0).getCell(i).getDateCellValue());
			}


		fis.close();

		return dateList;
	}

	//опреджеления наименования графика
	public static void getTitle() throws IOException {

			FileInputStream fis = new FileInputStream(fileName);
			Workbook wb = new HSSFWorkbook(fis);

			Graph.graphTitle = wb.getSheetAt(1).getRow(0).getCell(0).getStringCellValue();

			fis.close();
	}
}