package org.bubba.okcjugpoi;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.RadialGradientPaint;
import java.awt.Rectangle;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.title.TextTitle;
import org.jfree.chart.ui.HorizontalAlignment;
import org.jfree.chart.ui.RectangleEdge;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.graphics2d.svg.SVGGraphics2D;
import org.jfree.graphics2d.svg.SVGUtils;

public class SVGPieChart {

    private RadialGradientPaint createGradientPaint(Color c1, Color c2) {
        Point2D center = new Point2D.Float(0, 0);
        float radius = 200;
        float[] dist = {0.0f, 1.0f};
        return new RadialGradientPaint(center, radius, dist,
                new Color[]{c1, c2});
    }

    public void makeChart(XSSFWorkbook my_workbook) {
        XSSFSheet my_sheet = my_workbook.createSheet("SVG Pie Chart");

        JFreeChart chart = createChart(createDataset());
        SVGGraphics2D g2 = new SVGGraphics2D(600, 400);
        g2.setRenderingHint(JFreeChart.KEY_SUPPRESS_SHADOW_GENERATION, true);
        Rectangle r = new Rectangle(0, 0, 600, 400);
        chart.draw(g2, r);

        File f = new File("SVGPieChartDemo1.svg");
        try {
            SVGUtils.writeToSVG(f, g2.getSVGElement());
        } catch (IOException ex) {
            Logger.getLogger(SVGPieChart.class.getName()).log(Level.SEVERE, null, ex);
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

    private JFreeChart createChart(PieDataset dataset) {

        JFreeChart chart = ChartFactory.createPieChart(
                "Smart Phones Manufactured / Q3 2011", // chart title
                dataset);
        chart.removeLegend();

        // set a custom background for the chart
//        chart.setBackgroundPainter(new GradientPainter(new Color(20, 20, 20), 
//                RectangleAnchor.TOP_LEFT, Color.DARK_GRAY, 
//                RectangleAnchor.BOTTOM_RIGHT));
        // customise the title position and font
        TextTitle t = chart.getTitle();
        t.setHorizontalAlignment(HorizontalAlignment.LEFT);
        t.setPaint(new Color(240, 240, 240));
        t.setFont(new Font("Arial", Font.BOLD, 26));

        PiePlot plot = (PiePlot) chart.getPlot();
//        plot.setBackgroundPainter(null);
        plot.setInteriorGap(0.04);
//        plot.setBorderPainter(null);

        // use gradients and white borders for the section colours
        plot.setSectionPaint("Others", createGradientPaint(new Color(200, 200, 255), Color.BLUE));
        plot.setSectionPaint("Samsung", createGradientPaint(new Color(255, 200, 200), Color.RED));
        plot.setSectionPaint("Apple", createGradientPaint(new Color(200, 255, 200), Color.GREEN));
        plot.setSectionPaint("Nokia", createGradientPaint(new Color(200, 255, 200), Color.YELLOW));
        plot.setDefaultSectionOutlinePaint(Color.WHITE);
        plot.setSectionOutlinesVisible(true);
        plot.setDefaultSectionOutlineStroke(new BasicStroke(2.0f));

        // customise the section label appearance
        plot.setLabelFont(new Font("Courier New", Font.BOLD, 20));
        plot.setLabelLinkPaint(Color.WHITE);
        plot.setLabelLinkStroke(new BasicStroke(2.0f));
        plot.setLabelOutlineStroke(null);
        plot.setLabelPaint(Color.WHITE);
        plot.setLabelBackgroundPaint(null);
        // add a subtitle giving the data source
        TextTitle source = new TextTitle("Source: http://www.bbc.co.uk/news/business-15489523",
                new Font("Courier New", Font.PLAIN, 12));
        source.setPaint(Color.WHITE);
        source.setPosition(RectangleEdge.BOTTOM);
        source.setHorizontalAlignment(HorizontalAlignment.RIGHT);
        chart.addSubtitle(source);
        return chart;

    }

    private PieDataset createDataset() {
        DefaultPieDataset dataset = new DefaultPieDataset();
        dataset.setValue("Samsung", new Double(27.8));
        dataset.setValue("Others", new Double(55.3));
        dataset.setValue("Nokia", new Double(16.8));
        dataset.setValue("Apple", new Double(17.1));
        return dataset;
    }

}
