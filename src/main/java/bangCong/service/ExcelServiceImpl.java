package bangCong.service;

import bangCong.model.Employee;
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
            // tìm cột chứa tiêu đề ten nv, mã nv
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

}
