package org.formula;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author islam fawzy
 */
public class FormulaEditor {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {

        String input = "";
        String filename = "";

//      filename = "/media/islam/A2163849163820A91/java/formula-editor-v1.0/files/taxes.xls";
//      addSlabs(filename, 2, 2, "-s", "");
//      System.out.println(calculateTax(filename, "30600"));
        if (args[0].equalsIgnoreCase("-formula")) {
            input = args[1];
            try {
                String output = get_formula_result(input.trim());
                System.out.println(output);
            } catch (FormulaParseException ex) {
                System.out.println(ex.getMessage());
            }

        } else if (args[0].equalsIgnoreCase("-tax")) {
            input = args[1];
            filename = args[2];

            if (input.trim().length() == 0) {
                System.out.println("Error: No input");
                return;
            }
            if (!new File(filename).exists()) {
                System.out.println("Error: file not exist");
                return;
            }
            try {

                String output = calculateTax(filename, input.trim());
                System.out.println(output);
            } catch (FormulaParseException ex) {
                System.out.println(ex.getMessage());
            }

        } else if (args[0].equalsIgnoreCase("-slabs")) {

            if (args.length < 6) {
                System.out.println("Error: invalid argument");
                return;
            }

            int row = Integer.parseInt(args[1]);
            int cell = Integer.parseInt(args[2]);
            String type = args[3];
            String value = args[4];

            filename = args[5];

            addSlabs(filename, row, cell, type, value);
        }

    }

    public static void addSlabs(String filename, int row_num, int cell_num, String type, String value) {

        try {

            FileInputStream file = new FileInputStream(new File(filename));

            HSSFWorkbook workbook = new HSSFWorkbook(file);

            HSSFSheet sheet = workbook.getSheetAt(0);

            Cell cell = sheet.getRow(row_num).getCell(cell_num);

            if (type.equalsIgnoreCase("-n")) {         // numric value

                if (value.equalsIgnoreCase("-empty")) {
                    cell.setCellValue("");
                } else {
                    cell.setCellValue(Double.parseDouble(value));
                }

            } else if (type.equalsIgnoreCase("-s")) {   // if String

                if (value.equalsIgnoreCase("-empty")) {
                    cell.setCellValue("");
                } else {
                    cell.setCellValue(value);
                }

            } else if (type.equalsIgnoreCase("-f")) {   // if formula 

                if (value.equalsIgnoreCase("-empty")) {
                    cell.setCellValue("");
                } else {
                    cell.setCellFormula(value);
                }
            }

            /* write into the file */
            workbook = evaluateFormulas(workbook);
//          System.out.println(cell.getCellFormula());
//          System.out.println(cell.getStringCellValue());

            try {
                FileOutputStream out = new FileOutputStream(new File(filename));
                workbook.write(out);
                out.close();
            } catch (Exception ex) {
            }

        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }

    }

    public static String get_formula_result(String formula) throws FormulaParseException {

        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet("Payroll Formula");

        Row header = sheet.createRow(0);

        header.createCell(0).setCellValue("Payroll Formula");

        Row dataRow = sheet.createRow(1);

        // create formula result cell
        Cell formula_cell = dataRow.createCell(0);
        formula_cell.setCellType(Cell.CELL_TYPE_FORMULA);
        formula_cell.setCellFormula(formula);

        List<List> KpiLines = readWorkbook(workbook);

        String formula_result = KpiLines.get(1).get(0).toString();

        return formula_result;
    }

    private static String calculateTax(String filename, String value) {

        try {

            FileInputStream file = new FileInputStream(new File(filename));
            //Get the workbook instance for XLS file 
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            HSSFSheet sheet = workbook.getSheetAt(0);

            //Update the value of cell
            Cell cell = sheet.getRow(14).getCell(5);

            cell.setCellValue(Double.parseDouble(value));

            workbook = evaluateFormulas(workbook);

            Cell cell1 = sheet.getRow(14).getCell(6);
            double cell_value = cell1.getNumericCellValue();

            return String.valueOf(cell_value);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(FormulaEditor.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(FormulaEditor.class.getName()).log(Level.SEVERE, null, ex);
        }

        return "";
    }

    private static List<List> readWorkbook(HSSFWorkbook workbook) {

        List<List> lines = new ArrayList<List>();

        workbook = evaluateFormulas(workbook);

        HSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            List<String> line = new ArrayList<String>();

            //For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {

                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        line.add(new Boolean(cell.getBooleanCellValue()).toString());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
                            line.add(dateFormat.format(cell.getDateCellValue()));
                        } else {
                            line.add(new Double(cell.getNumericCellValue()).toString());
                        }
                        break;
                    case Cell.CELL_TYPE_STRING:
                        line.add(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        switch (cell.getCachedFormulaResultType()) {
                            case Cell.CELL_TYPE_NUMERIC:
                                line.add(new Double(cell.getNumericCellValue()).toString());
                                break;
                            case Cell.CELL_TYPE_STRING:
                                line.add(cell.getRichStringCellValue().toString());
                                break;
                        }
                        break;
                }
            }

            lines.add(line);
        }

        return lines;
    }

    private static HSSFWorkbook evaluateFormulas(HSSFWorkbook wb) {

        FormulaEvaluator evaluator = null;
        evaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = wb.getSheetAt(sheetNum);
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        evaluator.evaluateFormulaCell(c);
                        if (sheetNum == 0 && c.getColumnIndex() == r.getPhysicalNumberOfCells() - 1) {
                            switch (c.getCachedFormulaResultType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    break;
                            }
                        }
                    }
                }
            }
        }
        return wb;
    }

}
