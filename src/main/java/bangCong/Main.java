package bangCong;

import bangCong.service.ExcelService;
import bangCong.service.ExcelServiceImpl;

public class Main {
    public static void main(String[] args) {
        ExcelService excelService = new ExcelServiceImpl();
        excelService.readAndEmployee("./BangCong.xlsx");
    }
}