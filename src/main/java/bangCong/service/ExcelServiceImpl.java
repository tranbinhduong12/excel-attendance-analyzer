package bangCong.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class ExcelServiceImpl implements ExcelService {

    public Sheet readExcelFile(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        return workbook.getSheetAt(0);
    }

    public List<String> employeeInfo(Sheet sheet, String findInfo, List<String> infoEmployee) {
        int indexColumn = -1;

        for (Row row : sheet) {
            // tìm cột chứa tiêu đề tên nv, mã nv
            if(indexColumn == -1) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING &&
                            cell.getStringCellValue().trim().equalsIgnoreCase(findInfo.trim())) {
                        indexColumn = cell.getColumnIndex();
                        break;
                    }
                }
                continue; // sang dòng tiếp theo nếu tìm thấy tiêu đề
            }
            // Nếu đã tìm được cột, lấy dữ liệu ở các hàng dưới
            Cell targetCell = row.getCell(indexColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (targetCell.getCellType() == CellType.STRING) {
                String value = targetCell.getStringCellValue().trim();
                if (!value.isEmpty()) {
                    infoEmployee.add(value);
                }
            }
        }
        return infoEmployee;
    }

    public List<String> getWeekdayShifts(Sheet sheet, int startColumn, int endColumn ) {
        List<String> weekdayShifts = new ArrayList<>();

        for (Row row : sheet) {
            for (int i = startColumn; i <= endColumn; i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().startsWith("Tổng")) {
                    String shift = cell.getStringCellValue().replace("Tổng ", "").trim();
                    if (!shift.startsWith("WK")) {  // Loại bỏ ca chủ nhật
                        weekdayShifts.add(shift);
                    }
                }
            }
        }

        return weekdayShifts;
    }

    public List<String> getSundayShifts(Sheet sheet, int startColumn, int endColumn) {
        List<String> sundayShifts = new ArrayList<>();

        for (Row row : sheet) {
            for (int i = startColumn; i <= endColumn; i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().startsWith("Tổng")) {
                    String rawValue = cell.getStringCellValue().replace("Tổng ", "").trim();
                    if (rawValue.contains("&")) {
                        String[] parts = rawValue.split("&");
                        for (String part : parts) {
                            String shift = part.trim();
                            if (!shift.isEmpty()) {
                                sundayShifts.add(shift);
                            }
                        }
                    }
                }
            }
        }

        return sundayShifts;
    }


}
