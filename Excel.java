package ex;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Excel {

    public static void writeExcel(JTable data, String path) throws FileNotFoundException, IOException {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        /////
        int rownum = 0;
        Cell cell;
        Row row;
        /////
        HSSFCellStyle style = createStyleForTitle(workbook);
        row = sheet.createRow(rownum);
        /////
        TableModel model = data.getModel();
        for (int i = 0; i < model.getColumnCount(); i++) {
            cell = row.createCell(i, CellType.STRING);
            cell.setCellValue(model.getColumnName(i));
            cell.setCellStyle(style);
        }
        /////
        for (int i = 0; i < model.getRowCount(); i++) {
            rownum++;
            row = sheet.createRow(rownum);
            for (int j = 0; j < model.getColumnCount(); j++) {
                cell = row.createCell(j, CellType.STRING);
                cell.setCellValue(model.getValueAt(i, j).toString());
                cell.setCellStyle(style);
            }
        }
        /////
        File file = new File(path);
        file.getParentFile().mkdirs();
        FileOutputStream outFile = new FileOutputStream(file);
        workbook.write(outFile);
    }

    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {

        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;

    }

    public static DefaultTableModel readExcel(String path) throws FileNotFoundException, IOException {
        
        // Đọc một file XSL.
        FileInputStream inputStream = new FileInputStream(new File(path));

        // Đối tượng workbook cho file XSL.
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();
        DefaultTableModel dtm = null;
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                int hehe = 0;
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    hehe++;
                    Cell cell = cellIterator.next();
                }
                String header[] = new String[hehe];
                int i = 0;
                cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = dataFormatter.formatCellValue(cell);
                    header[i] = cellValue;
                    i++;
                }
                dtm = new DefaultTableModel(header, 0);
            } else {
                int hehe = 0;
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    hehe++;
                    Cell cell = cellIterator.next();
                }
                String header[] = new String[hehe];
                int i = 0;
                cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = dataFormatter.formatCellValue(cell);
                    header[i] = cellValue;
                    i++;
                }
                dtm.addRow(header);
                
            }

        }
//        for (int i = 0; i < dtm.getColumnCount(); i++) {
//            System.err.print(dtm.getColumnName(i)+"\t");
//        }
//        System.err.println("");
//        /////
//        int rownum = 0;
//        Row row;
//        for (int i = 0; i < dtm.getRowCount(); i++) {
//            rownum++;
//            row = sheet.createRow(rownum);
//            for (int j = 0; j < dtm.getColumnCount(); j++) {
//                System.err.print(dtm.getValueAt(i, j).toString()+"\t");
//                
//            }
//            System.err.println("");
//        }
        return dtm;
    }

}
