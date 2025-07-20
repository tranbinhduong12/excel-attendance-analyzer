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
        String path = "./BangCong.xlsx";
        File file = new File(path);
        List<String> allowedExtensions = Arrays.asList("xlsx", "xls");

        if (!ExcelUtils.isValidFile(file, allowedExtensions)) {
            System.out.println("invalid file");
            return;
        }

        ExcelService service = new ExcelServiceImpl();

        try {
            //  đọc file
            Sheet sheet = service.readExcelFile(file);
            System.out.println("Successfully read sheet: " + sheet.getSheetName());

            // lấy danh sách mã nv
            List<String> employeeID = service.employeeInfo(sheet, "Mã NV", new ArrayList<>());
            System.out.println("List of employee id: ");
            employeeID.forEach(System.out::println);

            // lấy danh sách tên nv
            List<String> employeeName = service.employeeInfo(sheet, "Họ tên", new ArrayList<>());
            System.out.println("\nList of employee name: ");
            employeeName.forEach(System.out::println);

            // danh sách ca ngày thường và ngày chủ nhật
            int startColumn = 3;
            int endColumn = 15;

            List<String> weekdayShifts = service.getWeekdayShifts(sheet, startColumn, endColumn);
            List<String> sundayShifts = service.getSundayShifts(sheet, startColumn, endColumn);

            int totalWeekdayShifts = weekdayShifts.size();
            int totalSundayShifts = sundayShifts.size();

            List<String> allShifts = new ArrayList<>();
            allShifts.addAll(weekdayShifts);
            allShifts.addAll(sundayShifts);

            System.out.println("\nList of weekday shifts:");
            weekdayShifts.forEach(System.out::println);

            System.out.println("\nList of sunday shifts:");
            sundayShifts.forEach(System.out::println);

        } catch (Exception e) {
            System.out.println("error reading excel file:" + e.getMessage());
        }
    }
}
