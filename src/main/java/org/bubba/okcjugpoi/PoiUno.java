package org.bubba.okcjugpoi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiUno {
    public static final String A__BUNCH__OF__DATA = "A Bunch Of Data";

    public void makeMeASpreadsheet() {

        //https://www.mkyong.com/java/apache-poi-reading-and-writing-excel-file-in-java/
        //HSSF is prefixed before the class name to indicate operations related to a Microsoft Excel 2003 file.
        //XSSF is prefixed before the class name to indicate operations related to a Microsoft Excel 2007 file or later.
        //XSSFWorkbook and HSSFWorkbook are classes which act as an Excel Workbook
        //HSSFSheet and XSSFSheet are classes which act as an Excel Worksheet
        //Row defines an Excel row
        //Cell defines an Excel cell addressed in reference to a row.
        
        XSSFWorkbook workbook = dosomethingfun();
        
        makeABunchOfData(workbook, A__BUNCH__OF__DATA);
        
        MakeBarChart makeBarChart = new MakeBarChart();
        makeBarChart.makeBarChartPlease(workbook);
//        makeBarChart.makeBarChartPlease2(workbook);
        
        MakeSomeFormulas makeSomeFormulas = new MakeSomeFormulas();
        makeSomeFormulas.makeSomeFormulasForMe(workbook);
        
//        DualAxisChart dualAxisChart = new DualAxisChart("Dual axis chart");
//        dualAxisChart.makeDualAxisChart(workbook);
         CreateSVGBarChart svg = new CreateSVGBarChart();
         svg.makeChart(workbook);
         
         SVGPieChart pie = new SVGPieChart();
         pie.makeChart(workbook);
        
        try {
            FileOutputStream outputStream = new FileOutputStream("File1.xls");
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private XSSFWorkbook dosomethingfun() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
        Object[][] datatypes = {
            {"xxxxxx Datatype xxxxxxx", "Type", "Size(in bytes)"},
            {"int", "Primitive", 2},
            {"float", "Primitive", 4},
            {"double", "Primitive", 8},
            {"char", "Primitive", 1},
            {"String", "Non-Primitive", "No fixed size"}
        };
        
//        sheet.setColumnWidth(0, 5000);
        sheet.setColumnWidth(1, 5000);
        sheet.setColumnWidth(2, 5000);

        int rowNum = 0;
        System.out.println("Creating excel");

        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {

                Cell cell = row.createCell(colNum++);

                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        return workbook;
    }

    private void makeABunchOfData(XSSFWorkbook workbook, String sheetName) {

        XSSFSheet sheet = workbook.createSheet(sheetName);
        workbook.setActiveSheet(workbook.getSheetIndex(sheetName));
        Random r = new Random();

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        for (int j = 0; j < 15; j++) {
            String x = String.valueOf((char) (j + 65));
            x += x;
            x += x;
            cell = row.createCell(j);
            cell.setCellValue(x);
            cell.setCellStyle(cellStyle);
        }

        for (int i = 1; i < 50; i++) {

            row = sheet.createRow(i);

            for (int j = 0; j < 15; j++) {
                cell = row.createCell(j);
                cell.setCellValue(new Random().nextInt(101 + 1));
                cell.setCellStyle(cellStyle);
            }
        }
        
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        
        row = sheet.createRow(100);
        cell = row.createCell(0);
        cell.setCellFormula("5*10*2");
        CellValue cellValue = evaluator.evaluate(cell);
        System.err.println("1 " + cellValue.getNumberValue());  //1 100.0
        
        row = sheet.createRow(101);
        cell = row.createCell(0);
        cell.setCellFormula("5*10*2");
        evaluator.evaluateFormulaCell(cell); 
        System.out.println("2 " + cell.getCellType());     //2 FORMULA
        System.out.println("2 " + cell.getCellFormula());  //2 5*10*2
        
        row = sheet.createRow(102);
        cell = row.createCell(0);
        cell.setCellFormula("5*10*2");
        evaluator.evaluateInCell(cell); 
        System.out.println("3 " + cell.getCellType());          //3 NUMERIC
        System.out.println("3 " + cell.getNumericCellValue());  //3 100.0
    }
}
