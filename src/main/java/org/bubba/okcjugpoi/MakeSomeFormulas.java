package org.bubba.okcjugpoi;

import java.math.BigInteger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MakeSomeFormulas {

    public void makeSomeFormulasForMe(XSSFWorkbook my_workbook) {

        XSSFSheet dataSheet = my_workbook.getSheet(PoiUno.A__BUNCH__OF__DATA);
        
        XSSFRow row;
        XSSFCell cell;
        BigInteger bi = BigInteger.ZERO;
        dataSheet.getRow(0).createCell(15).setCellValue("ToTaLs");
        
        for(int i = 1; i < 50; i++){
            for(int j = 0; j < 15; j++){
                row = dataSheet.getRow(i);
                cell = row.getCell(j);
                System.err.println("aaa " + i + " " + j + " " + cell.getNumericCellValue() + "--");
                String theNbr = Double.toString(cell.getNumericCellValue());
                theNbr = theNbr.substring(0, theNbr.length()-2);
                bi = bi.add(new BigInteger(theNbr));
            }
            dataSheet.getRow(i).createCell(15).setCellFormula("SUM(A" + (i + 1) +":O"+ (i + 1) + ")");
            dataSheet.getRow(i).createCell(16).setCellFormula(bi.toString());
        }
        dataSheet.createRow(50);
        
        for(int i = 0; i < 15; i++){
//            for(int j = 0; j < 50; j++){
//                row = dataSheet.getRow(i);
//                cell = row.getCell(j);
//                System.err.println("aaa " + i + " " + j + " " + cell.getNumericCellValue() + "--");
//                String theNbr = Double.toString(cell.getNumericCellValue());
//                theNbr = theNbr.substring(0, theNbr.length()-2);
//                bi = bi.add(new BigInteger(theNbr));
//            }
            String x = String.valueOf((char) (i + 65));
            dataSheet.getRow(50).createCell(i).setCellFormula("SUM(" + x + "2:" + x + "50)");
        }
    }
}
