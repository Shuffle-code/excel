package site.example.ecxel.parser;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.richParser.XMLStreamReaderExtImpl;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
@Slf4j
@Component
public class ExcelParser {


    @Value("${storage.location}")
    private Path storagePath;
    public  void readXlsx(){
        readXlsxFile("C:\\Users\\79130\\Documents\\java\\test.xlsx");
    }

    public void readXlsxFile(String pathFile) {
        try {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(pathFile));
            XSSFSheet xssfSheet = xssfWorkbook.getSheet("Лист1");
//            XSSFRow row = xssfSheet.getRow(1);
            XSSFRow row;
            int i = 1;
            while ((row = xssfSheet.getRow(i)) != null){
                try {
                    System.out.println("№ = " + row.getCell(4).getNumericCellValue());
                }catch (Exception е) {
                    log.error("Ошибка в стоке № : " + row.getCell(4).getColumnIndex());
                }
                System.out.println("№ = " + row.getCell(0).getNumericCellValue());
//                System.out.println("? = " + row.getCell(4).getDateCellValue());
                i++;
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void printCellValue(Cell cell) {
        CellType cellType = cell.getCellType().equals(CellType.FORMULA)
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (cellType.equals(CellType.STRING)) {
            System.out.print(cell.getStringCellValue() + " | ");
        }
        if (cellType.equals(CellType.NUMERIC)) {
            if (DateUtil.isCellDateFormatted(cell)) {
                System.out.print(cell.getDateCellValue() + " | ");
            } else {
                System.out.print(cell.getNumericCellValue() + " | ");
            }
        }
        if (cellType.equals(CellType.BOOLEAN)) {
            System.out.print(cell.getBooleanCellValue() + " | ");
        }
    }
}
