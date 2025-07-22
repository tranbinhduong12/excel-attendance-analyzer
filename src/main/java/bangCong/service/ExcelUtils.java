package bangCong.service;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.util.List;

public class ExcelUtils {
    // kiểm tra file đúng định dạng
    public static boolean isValidFile(File file, List<String> allowedExtensions) {
        if (file == null) return false;
        String name = file.getName().toLowerCase();
        return allowedExtensions.stream()
                .map(String::toLowerCase)
                .anyMatch(ext -> name.endsWith("." + ext));
    }

    // Kiểm tra 1 dòng có trống hay không (không có ô nào chứa dữ liệu)
    public static boolean isRowEmpty(Row row) {
        if (row == null) return true;

        for (Cell cell : row) {
            if (cell.getCellType() != CellType.BLANK) {
                // Nếu có 1 ô không rỗng → không phải dòng trống
                return false;
            }
        }
        return true;
    }

    // Kiểm tra dòng có phải là dòng chứa dữ liệu nhân viên hợp lệ không
    public static boolean isValidEmployeeRow(Row row) {
        if (row == null) return false;

        Cell cellMaNV = row.getCell(1); // Cột "Mã NV"
        Cell cellTenNV = row.getCell(2); // Cột "Họ Tên"

        if (cellMaNV == null || cellTenNV == null) return false;

        boolean hasMaNV = cellMaNV.getCellType() == CellType.STRING && !cellMaNV.getStringCellValue().trim().isEmpty();
        boolean hasTenNV = cellTenNV.getCellType() == CellType.STRING && !cellTenNV.getStringCellValue().trim().isEmpty();

        return hasMaNV && hasTenNV;
    }

    // lấy giá trị String
    public static String getCellString(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return "";

        if (cell.getCellType() == CellType.FORMULA) {
            CellValue value = evaluator.evaluate(cell);
            if (value != null) {
                switch (value.getCellType()) {
                    case STRING: return value.getStringValue();
                    case NUMERIC: return String.valueOf(value.getNumberValue());
                    case BOOLEAN: return String.valueOf(value.getBooleanValue());
                }
            }
        } else {
            return cell.toString().trim();
        }
        return "";
    }

    // lấy giá trị số
    public static double getCellNumeric(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return 0;

        CellType cellType = cell.getCellType();

        if (cellType == CellType.FORMULA) {
            CellValue cellValue = evaluator.evaluate(cell);
            if (cellValue != null && cellValue.getCellType() == CellType.NUMERIC) {
                return cellValue.getNumberValue();
            }
        } else if (cellType == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        return 0;
    }

    public static boolean isSunday(Sheet sheet, int day) {
        Row row = sheet.getRow(5); // Dòng tiêu đề ca
        if (row == null) return false;

        int colsPerDay = 5;
        int startCol = 16 + (day - 1) * colsPerDay;

        // Kiểm tra cột đầu tiên trong nhóm 5 cột của ngày đó
        for (int i = 0; i < 5; i++) {
            Cell cell = row.getCell(startCol + i);
            if (cell != null) {
                String text = cell.toString().toUpperCase();
                if (text.contains("WK")){
                    System.out.println("........" + i + text);
                    return true;
                }
            }
        }
        return false;
    }

}