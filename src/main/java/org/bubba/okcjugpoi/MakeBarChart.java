package org.bubba.okcjugpoi;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

public class MakeBarChart {

    private static final String FIRST_DATA = "firstData";

    private void makeSomeMoreData(XSSFWorkbook my_workbook) {

        XSSFSheet sheet = my_workbook.createSheet(FIRST_DATA);
        Random r = new Random();

        for (int j = 0; j < 10; j++) {
            Row row = sheet.createRow(j);
            Cell cell = row.createCell(0);

            String x = getMeALabel(j);
            cell = row.createCell(0);
            cell.setCellValue(x);

            CellStyle cellStyle = cell.getCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle = cell.getCellStyle();
            cell.setCellStyle(cellStyle);
// ----------
            cell = row.createCell(1);
            cell.setCellValue(new Random().nextInt(101 + 1));
            cell.setCellStyle(cellStyle);
        }
    }

    public void makeBarChartPlease(XSSFWorkbook my_workbook) {

        makeSomeMoreData(my_workbook);

        XSSFSheet my_sheet = my_workbook.createSheet("Bar Chart Sheet");
        XSSFSheet dataSheet = my_workbook.getSheet(FIRST_DATA);
        /* Create Dataset that will take the chart data */
        DefaultCategoryDataset my_bar_chart_dataset = new DefaultCategoryDataset();
        /* We have to load bar chart data now */
 /* Begin by iterating over the worksheet*/
 /* Create an Iterator object */
        Iterator<Row> rowIterator = dataSheet.iterator();
        /* Loop through worksheet data and populate bar chart dataset */
        String chart_label = "a";
        Number chart_data = 0;

        while (rowIterator.hasNext()) {
            //Read Rows from Excel document
            Row row = rowIterator.next();
            //Read cells in Rows and get chart data
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                if (cell.getCellType().equals(CellType.NUMERIC)) {
                    chart_data = cell.getNumericCellValue();
                } else if (cell.getCellType().equals(CellType.STRING)) {
                    chart_label = cell.getStringCellValue();
                }
//                System.err.println("aaaaa" + chart_label + " " + chart_data);
            }
            /* Add data to the data set */
 /* We don't have grouping in the bar chart, so we put them in fixed group */
            my_bar_chart_dataset.addValue(chart_data.doubleValue(), "Marks", chart_label);
        }

        /* Create a logical chart object with the chart data collected */
        JFreeChart barChartObject = ChartFactory.createBarChart("Subject Vs Marks", "Subject", "Marks", my_bar_chart_dataset, PlotOrientation.VERTICAL, true, true, false);
        /* Dimensions of the bar chart */
        int width = 640;
        /* Width of the chart */
        int height = 480;
        /* Height of the chart */
 /* We don't want to create an intermediate file. So, we create a byte array output stream 
                and byte array input stream
                And we pass the chart data directly to input stream through this */
 /* Write chart as PNG to Output Stream */
        ByteArrayOutputStream chart_out = new ByteArrayOutputStream();
        byte[] imageInByte = null;
        try {
//            ChartUtilities.writeChartAsPNG(chart_out, barChartObject, width, height);
            BufferedImage bi = barChartObject.createBufferedImage(width, height);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(bi, "png", baos);
            baos.flush();
            imageInByte = baos.toByteArray();
            baos.close();
        } catch (IOException ex) {
            Logger.getLogger(MakeBarChart.class.getName()).log(Level.SEVERE, null, ex);
        }

        /* We can now read the byte data from output stream and stamp the chart to Excel worksheet */
//        int my_picture_id = my_workbook.addPicture(chart_out.toByteArray(), Workbook.PICTURE_/TYPE_PNG);
        int my_picture_id = my_workbook.addPicture(imageInByte, Workbook.PICTURE_TYPE_PNG);

        try {
            /* we close the output stream as we don't need this anymore */
            chart_out.close();
        } catch (IOException ex) {
            Logger.getLogger(MakeBarChart.class.getName()).log(Level.SEVERE, null, ex);
        }

        /* Create the drawing container */
        XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
        /* Create an anchor point */
        ClientAnchor my_anchor = new XSSFClientAnchor();
        /* Define top left corner, and we can resize picture suitable from there */
        my_anchor.setCol1(4);
        my_anchor.setRow1(5);
        /* Invoke createPicture and pass the anchor point and ID */
        XSSFPicture my_picture = drawing.createPicture(my_anchor, my_picture_id);
        /* Call resize method, which resizes the image */
        my_picture.resize();
    }

//    public void makeBarChartPlease2(XSSFWorkbook my_workbook) {
//
//        XSSFSheet my_sheet = my_workbook.createSheet("Bar Chart Sheet Dos");
//        XSSFSheet dataSheet = my_workbook.getSheet(FIRST_DATA);
//        /* Create Dataset that will take the chart data */
//        DefaultCategoryDataset my_bar_chart_dataset = new DefaultCategoryDataset();
//
//        /* We have to load bar chart data now */
// /* Begin by iterating over the worksheet*/
// /* Create an Iterator object */
//        Iterator<Row> rowIterator = dataSheet.iterator();
//        /* Loop through worksheet data and populate bar chart dataset */
//        String chart_label = "a";
//        Number chart_data = 0;
//
//        while (rowIterator.hasNext()) {
//            //Read Rows from Excel document
//            Row row = rowIterator.next();
//            //Read cells in Rows and get chart data
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//
//                if (cell.getCellType().equals(CellType.NUMERIC)) {
//                    chart_data = cell.getNumericCellValue();
//                } else if (cell.getCellType().equals(CellType.STRING)) {
//                    chart_label = cell.getStringCellValue();
//                }
//            }
//            /* Add data to the data set */
// /* We don't have grouping in the bar chart, so we put them in fixed group */
//            CategoryAxis categoryAxis = new CategoryAxis((chart_label + " " + chart_data.doubleValue()));
//            categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);
//            categoryAxis.setLabel(chart_label + " " + chart_data);
//
//            my_bar_chart_dataset.addValue(chart_data.doubleValue(), "Marks", (chart_label + " " + chart_data.doubleValue()));
//        }
//
//        /* Create a logical chart object with the chart data collected */
//        JFreeChart barChartObject = ChartFactory.createBarChart("Subject Vs Marks", "Subject", "Marks", my_bar_chart_dataset, PlotOrientation.VERTICAL, true, true, false);
//        barChartObject.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions.DOWN_45); //http://www.java2s.com/Code/Java/Chart/JFreeChartDualAxisDemo.htm
////        barChartObject.getCategoryPlot().getDomainAxis().setCategoryLabelPositions(CategoryLabelPositions.UP_45); //http://www.java2s.com/Code/Java/Chart/JFreeChartDualAxisDemo.htm
//
//        int width = 640;
//        int height = 480;
//
//
//        /* We don't want to create an intermediate file. So, we create a byte array output stream 
//                and byte array input stream
//                And we pass the chart data directly to input stream through this */
// /* Write chart as PNG to Output Stream */
//        ByteArrayOutputStream chart_out = new ByteArrayOutputStream();
//
////        try {
////            ChartUtilities.writeChartAsPNG(chart_out, barChartObject, width, height);
////        } catch (IOException ex) {
////            Logger.getLogger(MakeBarChart.class.getName()).log(Level.SEVERE, null, ex);
////        }
//
//        /* We can now read the byte data from output stream and stamp the chart to Excel worksheet */
//        int my_picture_id = my_workbook.addPicture(chart_out.toByteArray(), Workbook.PICTURE_TYPE_PNG);
//
//        try {
//            /* we close the output stream as we don't need this anymore */
//            chart_out.close();
//        } catch (IOException ex) {
//            Logger.getLogger(MakeBarChart.class.getName()).log(Level.SEVERE, null, ex);
//        }
//
//        /* Create the drawing container */
//        XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
//        /* Create an anchor point */
//        ClientAnchor my_anchor = new XSSFClientAnchor();
//        /* Define top left corner, and we can resize picture suitable from there */
//        my_anchor.setCol1(4);
//        my_anchor.setRow1(5);
//        /* Invoke createPicture and pass the anchor point and ID */
//        XSSFPicture my_picture = drawing.createPicture(my_anchor, my_picture_id);
//        /* Call resize method, which resizes the image */
//        my_picture.resize();
//    }

    private String getMeALabel(int j) {
        String x = String.valueOf((char) (j + 65));
        x += x;
        x += x;
        return x;
    }
}
