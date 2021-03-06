package org.bubba.okcjugpoi;

import java.awt.Color;
import java.awt.GradientPaint;
import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.renderer.category.StatisticalBarRenderer;
import org.jfree.chart.ui.TextAnchor;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.statistics.DefaultStatisticalCategoryDataset;
import org.jfree.graphics2d.svg.SVGGraphics2D;
import org.jfree.graphics2d.svg.SVGUtils;

public class CreateSVGBarChart {

    public void makeChart(XSSFWorkbook my_workbook) {
        
        XSSFSheet dataSheet = my_workbook.getSheet("A Bunch Of Data");
        XSSFSheet my_sheet = my_workbook.createSheet("SVG Bar Chart");
        
        JFreeChart chart = createChart(createDataset(dataSheet));
        
        SVGGraphics2D g2 = new SVGGraphics2D(600, 400);
        Rectangle r = new Rectangle(0, 0, 600, 400);
        chart.draw(g2, r);
        
        File f = new File("SVGBarChartDemo1.svg");
        
        try {
            SVGUtils.writeToSVG(f, g2.getSVGElement());
        } catch (IOException ex) {
            Logger.getLogger(CreateSVGBarChart.class.getName()).log(Level.SEVERE, null, ex);
        }

        int width = 640;
        int height = 480;

        byte[] imageInByte = null;

        try {
            BufferedImage bi = chart.createBufferedImage(width, height);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(bi, "png", baos);
            baos.flush();
            imageInByte = baos.toByteArray();
            baos.close();
        } catch (IOException ex) {
            Logger.getLogger(MakeBarChart.class.getName()).log(Level.SEVERE, null, ex);
        }
        int my_picture_id = my_workbook.addPicture(imageInByte, Workbook.PICTURE_TYPE_PNG);

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

    private CategoryDataset createDataset(XSSFSheet sheet) {
        DefaultStatisticalCategoryDataset dataset
                = new DefaultStatisticalCategoryDataset();

        for (int i = 0; i < 5; i++) {

            XSSFCell columnHeading = sheet.getRow(0).getCell(i);

            XSSFCell cell1 = sheet.getRow(1).getCell(i);
            XSSFCell cell2 = sheet.getRow(2).getCell(i);

            dataset.add(
                    cell1.getNumericCellValue(),
                    cell2.getNumericCellValue(),
                    "Row 1",
                    columnHeading.getStringCellValue());
            
            cell1 = sheet.getRow(3).getCell(i);
            cell2 = sheet.getRow(4).getCell(i);

            dataset.add(
                    cell1.getNumericCellValue(),
                    cell2.getNumericCellValue(),
                    "Row 2",
                    columnHeading.getStringCellValue());

        }

        return dataset;
    }

    private JFreeChart createChart(CategoryDataset dataset) {

        JFreeChart chart = ChartFactory.createLineChart(
                "Statistical Bar Chart Demo 1", "Type", "Value", dataset);

        CategoryPlot plot = (CategoryPlot) chart.getPlot();

        // customise the range axis...
        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
        rangeAxis.setAutoRangeIncludesZero(false);

        // customise the renderer...
        StatisticalBarRenderer renderer = new StatisticalBarRenderer();
        renderer.setDrawBarOutline(false);
        renderer.setErrorIndicatorPaint(Color.black);
        renderer.setIncludeBaseInRange(false);
        plot.setRenderer(renderer);

        // ensure the current theme is applied to the renderer just added
        ChartUtils.applyCurrentTheme(chart);

        renderer.setDefaultItemLabelGenerator(
                new StandardCategoryItemLabelGenerator());
        renderer.setDefaultItemLabelsVisible(true);
        renderer.setDefaultItemLabelPaint(Color.yellow);
        renderer.setDefaultPositiveItemLabelPosition(new ItemLabelPosition(
                ItemLabelAnchor.INSIDE6, TextAnchor.BOTTOM_CENTER));

        // set up gradient paints for series...
        GradientPaint gp0 = new GradientPaint(0.0f, 0.0f, Color.blue,
                0.0f, 0.0f, new Color(0, 0, 64));
        GradientPaint gp1 = new GradientPaint(0.0f, 0.0f, Color.green,
                0.0f, 0.0f, new Color(0, 64, 0));
        renderer.setSeriesPaint(0, gp0);
        renderer.setSeriesPaint(1, gp1);
        return chart;
    }

}
