package manipulando_diretorio_java;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main_Lendo_Arquivo_Excel {

    public static void main(String[] args) throws IOException {
        LendoExcel();
    }
    
    public static void LendoExcel(){
         try {
             // no Projeto temos um arquivo modelo para lermos, deve salvar o arquivo na rota e apontar no path
            FileInputStream file = new FileInputStream(new File("C:\\EducaCienciaFastCode\\ModeloLeituraExcel.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (IOException e) {
            System.out.println(e);
        }
    }
}
