import com.aspose.pdf.*;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// export some data from pdf file to excel program.

public class pdf_to_excel {

    public static void main(String[] args) {
        String pdf_file_path = "data/test.pdf";
        String output_file_path = "data/test.xlsx";

        // convert whole pdf to excel file.
        String temp_output_file = "temp.xlsx";
        pdf_to_excel(pdf_file_path, temp_output_file);

        // read the table
        List<List<String>> table = excel_reader(temp_output_file);

        // write this table to excel
        excel_writer(table, output_file_path);

        // delete temporary file
        delete_file(temp_output_file);
    }

    public static boolean isNumericOrNull(String str){
        return str == null || str.isEmpty() || str.matches("[0-9.]+");
    }

    static void delete_file(String file_path) {
        File file = new File(file_path);
        try {
            file.delete();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void pdf_to_excel(String pdf_file_path, String excel_file_path) {
        Path path = Paths.get(pdf_file_path);
        Document doc = new Document(path.toString());

        // Set Excel options
        ExcelSaveOptions options = new ExcelSaveOptions();
        // Set output format
        options.setFormat(ExcelSaveOptions.ExcelFormat.XLSX);
        // Convert PDF to XLSX
        doc.save(excel_file_path, options);
    }

    static List<List<String>> excel_reader(String excel_file_path) {

        List<List<String>> table = new ArrayList<List<String>>();

        try {
            File file = new File(excel_file_path);
            FileInputStream fis = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> row_itr = sheet.iterator();
            int valid_row = 0;
            boolean is_table = false;

            while (row_itr.hasNext())
            {
                if (is_table == true) {
                    valid_row++;
                }
                Row row = row_itr.next();
                Iterator<Cell> cell_itr = row.cellIterator();
                int cell_index = 0;
                List<String> row_data = new ArrayList<String>();
                while (cell_itr.hasNext())
                {
                    Cell cell = cell_itr.next();
                    cell_index++;

                    // TODO: check with all other patterns of STT
                    if (is_table == false && cell.getStringCellValue().equals("STT")) {
                        System.out.println("table detected.");
                        is_table = true;
                    }
                    if (is_table == false) break;

                    int cell_type = cell.getCellType();
                    // ignore check for first row and second row (valid_row > 1).
                    if (valid_row > 1 && cell_index == 1 && cell_type == Cell.CELL_TYPE_STRING) {
                        if (!isNumericOrNull(cell.getStringCellValue())) {
                            System.out.println("End of table.");
                            System.out.format("valid row %d%n", valid_row);
                            System.out.println("content: " + cell.getStringCellValue());
                            is_table = false;
                            break;
                        }
                    }

                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_STRING:
                            row_data.add(cell.getStringCellValue());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            row_data.add(String.valueOf(cell.getNumericCellValue()));
                            break;
                        default:
                            System.out.format("Unknown cell type. row %d, cell %d%n", valid_row, cell_index);
                    }
                }
                if (is_table == true) {
                    table.add(row_data);
                }
            }
            fis.close();
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }

        return table;
    }

    static void excel_writer(List<List<String>> table, String excel_file_path) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("sheet 1");
        for (int row_index = 0; row_index < table.size(); row_index++) {
            Row row = spreadsheet.createRow(row_index);
            for (int cell_index = 0; cell_index < table.get(row_index).size(); cell_index++) {
                Cell cell = row.createCell(cell_index);
                cell.setCellValue(table.get(row_index).get(cell_index));
            }
        }
        try {
            FileOutputStream out = new FileOutputStream(new File(excel_file_path));
            workbook.write(out);
            out.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
}