package bangCong;

import bangCong.service.ExcelService;
import bangCong.service.ExcelServiceImpl;
import bangCong.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String path = "./BangCong.xlsx";
        File file = new File(path);
        List<String> allowedExtensions = Arrays.asList("xlsx", "xls");

        if (!ExcelUtils.isValidFile(file, allowedExtensions)) {
            System.out.println("invalid file");
            return;
        }

        ExcelService service = new ExcelServiceImpl();
        try {
            Sheet sheet = service.readExcelFile(file);
            System.out.println("Successfully read sheet: " + sheet.getSheetName());
        } catch (Exception e) {
            System.out.println("error reading excel file:" + e.getMessage());
        }
    }
}
