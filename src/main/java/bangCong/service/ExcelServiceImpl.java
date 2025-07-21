package bangCong.service;

import bangCong.model.Employee;
import bangCong.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static bangCong.utils.ExcelUtils.getCellNumeric;
import static bangCong.utils.ExcelUtils.getCellString;


public class ExcelServiceImpl implements ExcelService {

    public void readAndEmployee(String filePath) {
        try(FileInputStream fis = new FileInputStream(new File(filePath));
            Workbook workbook = WorkbookFactory.create(fis)) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Sheet sheet = workbook.getSheetAt(0);
            List<Employee> employeeList = new ArrayList<>();

//            Row headerRow = sheet.getRow(0); // check tiêu đề
//            for (Cell cell : headerRow) {
//                System.out.println("Cột " + cell.getColumnIndex() + ": " + cell.toString());
//            }

            for (int i = 4; i <= sheet.getLastRowNum();i++) { // bỏ dòng tiêu đề

                Row row = sheet.getRow(i);
                if (row == null) continue;

                if (ExcelUtils.isRowEmpty(row)) continue; // Bỏ dòng trống
                if (!ExcelUtils.isValidEmployeeRow(row)) continue; // Bỏ dòng không phải nhân viên

                String maNV = getCellString(row.getCell(1), evaluator);
                String tenNV = getCellString(row.getCell(2), evaluator);

                // Giờ làm từ các cột CN, GC, TC, GC1, TC1, WK-D, WK-N (index 3,5,7,9,11,13,14)
                double cn = getCellNumeric(row.getCell(3), evaluator);
                double gc = getCellNumeric(row.getCell(5), evaluator);
                double tc = getCellNumeric(row.getCell(7), evaluator);
                double gc1 = getCellNumeric(row.getCell(9), evaluator);
                double tc1 = getCellNumeric(row.getCell(11), evaluator);
                double wkD = getCellNumeric(row.getCell(13), evaluator);
                double wkN = getCellNumeric(row.getCell(14), evaluator);
//                double totalHours = wkN;

                double totalHours = cn + gc + tc + gc1 + tc1 + wkD + wkN;

                // Tiền từ các cột $CN, $GC, $TC, $GC1, $TC1, $WK (index 4,6,8,10,12,15)
                double moneyCN = getCellNumeric(row.getCell(4), evaluator);
                double moneyGC = getCellNumeric(row.getCell(6), evaluator);
                double moneyTC = getCellNumeric(row.getCell(8), evaluator);
                double moneyGC1 = getCellNumeric(row.getCell(10), evaluator);
                double moneyTC1 = getCellNumeric(row.getCell(12), evaluator);
                double moneyWK = getCellNumeric(row.getCell(15), evaluator);
//
                double totalMoney = cn*moneyCN + gc*moneyGC + tc*moneyTC + gc1*moneyGC1 + tc1*moneyTC1 + (wkD+wkN)*moneyWK;
//                double totalMoney = moneyGC1;

//                System.out.println("Row " + i + ": CN = " + cn + ", GC = " + gc + ", TC = " + tc);
//                System.out.println("Money CN = " + moneyCN + ", GC = " + moneyGC + ", TC = " + moneyTC);

                // in ra tổng tiền cột Q
                double colQ = getCellNumeric(row.getCell(16), evaluator);

                double actualTotalFromExcel = getCellNumeric(row.getCell(16), evaluator);




                Employee emp = new Employee(maNV, tenNV);
                emp.setTotalHoursWorked(totalHours);
                emp.setTotalEaring(totalMoney);
                emp.setActualTotalFromExcel(actualTotalFromExcel);
                employeeList.add(emp);
            }
            //in kết quả
            for(Employee emp : employeeList) {
                System.out.println(emp);

            }
        }catch (IOException e) {
            System.out.println("Lỗi đọc file " + e.getMessage());
        }
    }

}