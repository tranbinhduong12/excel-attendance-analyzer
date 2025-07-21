package bangCong;

import bangCong.service.ExcelService;
import bangCong.service.ExcelServiceImpl;
import bangCong.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        ExcelService excelService = new ExcelServiceImpl();
        excelService.readAndEmployee("./BangCong.xlsx");
    }
}