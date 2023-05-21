package site.example.ecxel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;
import site.example.ecxel.parser.ExcelParser;

@SpringBootApplication
public class EcxelApplication {



    public static void main(String[] args) {
        ConfigurableApplicationContext context = SpringApplication.run(EcxelApplication.class, args);
        ExcelParser excelParser = context.getBean(ExcelParser.class);
        excelParser.readXlsx();
    }

}
