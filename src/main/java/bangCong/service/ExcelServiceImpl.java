package bangCong.service;

import bangCong.model.Employee;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import static bangCong.service.ExcelUtils.getCellNumeric;
import static bangCong.service.ExcelUtils.getCellString;


public class ExcelServiceImpl implements ExcelService {

    public List<String> listCaNgayThuong(Sheet sheet, int startIxColumn, int endIxColumn) {
        List<String> listTenCas = new ArrayList<>();
        for (Row row : sheet) { // duyệt các cell trong sheet
            for(int i = startIxColumn;i<=endIxColumn;i++) { // duyệt các cell trong khoảng startCol đến endCol
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                // Tìm các cell bắt đầu bằng "Tổng" nhưng k chứa "WK"
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().startsWith("Tổng")) {
                    String ca = cell.getStringCellValue().replace("Tổng ", "").trim();
                    if(!ca.startsWith("WK")){
                        listTenCas.add(ca); // lưu tên ca vào ds
                    }
                }
            }
        }
        return listTenCas;
    }

    // giống listCaNgayThuong nhưng tìm các cell bắt đầu bằng "Tổng" và có "&", tách chuỗi bằng "&" để lấy danh sách ca
    public List<String> listCaDaySunday(Sheet sheet,int startIxColumn,int endIxColumn) {
        List<String> listCaDaySunday = new ArrayList<>();
        for (Row row : sheet) {
            for(int i = startIxColumn;i<=endIxColumn;i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().startsWith("Tổng")) {
                    String ca = cell.getStringCellValue().replace("Tổng ", "").trim();
                    if(ca.contains("&")) {
                        String[] parts = ca.split("&");
                        for (String part : parts) {
                            part = part.trim();
                            listCaDaySunday.add(part);
                        }
                    }
                }
            }
        }
        return listCaDaySunday;
    }

    // tìm cột chứa tiêu đề findInfo -> lấy giá trị chuỗi từ dòng tiếp theo
    public List<String> extractEmployeeInfo(Sheet sheet,String findInfo,List<String> infoNV) {
        int indexColumn = -1;
        for (Row row : sheet) {
            if(indexColumn==-1){
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(findInfo)) {
                        indexColumn = cell.getColumnIndex();
                        break;
                    }
                }
                continue;
            }
            Cell cellTypeHourWork = row.getCell(indexColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

            if (cellTypeHourWork.getCellType() == CellType.STRING ) {
                infoNV.add(cellTypeHourWork.getStringCellValue());
            }
        }
        return infoNV;
    }

    // tìm cột chứa tiêu đề ca -> bỏ qua dòng này
    public List<Double> totalHourseWork(Sheet sheet,int indexStarRow,String findInfo,int indexEndRow) {
        int indexColumn = -1;
        List<Double> hourseWork = new ArrayList<>();
        for (Row row : sheet) {
            if (indexColumn == -1) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(findInfo)) {
                        indexColumn = cell.getColumnIndex();
                        break;
                    }
                }
                continue;
            }
        }
        //lấy giá trị từ các dòng trong khoảng indexStarRow đến indexEndRow
        for(int i=indexStarRow;i<indexEndRow;i++){
            Row rows = sheet.getRow(i);
            Cell cellTypeHourWork = rows.getCell(indexColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if(cellTypeHourWork.getCellType() != CellType.FORMULA){
                hourseWork.add(0.0); // nếu cell k có giá trị trả về 0.0
            }
            if(cellTypeHourWork.getCellType() == CellType.FORMULA){
                hourseWork.add(cellTypeHourWork.getNumericCellValue());
            }
            if(i==indexEndRow-1){
                break;
            }
        }
        return hourseWork;
    }

    // giống totalHoursWork, tìm cột chứa tiêu đề lương, bỏ qua dòng này
    public List<Double> totalPriceCa(Sheet sheet,int indexStarRow,String findInfo,int indexEndRow) {
        int indexColumn = -1;
        List<Double> totalPriceCa = new ArrayList<>();
        for (Row row : sheet) {
            if (indexColumn == -1) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(findInfo)) {
                        if(findInfo.startsWith("WK")) { // nếu findInfo bắt đầu bằng "WK", lấy giá trị từ cột 14
                            indexColumn = 14;
                            break;
                        }
                        indexColumn = cell.getColumnIndex();
                        break;
                    }
                }
                continue;
            }
        }
        //lấy giá trị từ các dòng trong khoảng (indexStarRow+1) đến indexEndRow
        for(int i=indexStarRow;i<indexEndRow;i++){
            Row rows = sheet.getRow(i);
            Cell cellTypeHourWork = rows.getCell(indexColumn+1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if(cellTypeHourWork.getCellType() != CellType.FORMULA){
                totalPriceCa.add(0.0);
            }
            if(cellTypeHourWork.getCellType() == CellType.FORMULA){
                totalPriceCa.add(cellTypeHourWork.getNumericCellValue());
            }
        }
        return totalPriceCa;
    }

    // lấy số giờ làm việc của một nhân viên trong khoảng cột từ indexStartColumn đến indexEndColumn tại dòng indexRow.
    public List<Double> countWorkingDaysByDate(Sheet sheet,int indexStartColumn,int indexRow,int indexEndColumn) {
        List<Double> workingDaysByDate = new ArrayList<>();
        for (Row row : sheet) {
            if(sheet.getRow(indexRow) == null){
                System.out.println("Row is null");
                return null;
            }
            if(row == sheet.getRow(indexRow)){
                for(int indexColumn=indexStartColumn;indexColumn<=indexEndColumn;indexColumn++){
                    Cell cellTypeHourWork = sheet.getRow(indexRow).getCell(indexColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if(cellTypeHourWork.getColumnIndex()<=indexEndColumn){
                        if (cellTypeHourWork.getCellType() == CellType.NUMERIC ) { // nếu là ô số -> thêm vào list
                            workingDaysByDate.add(cellTypeHourWork.getNumericCellValue());
                        }else if (cellTypeHourWork.getCellType() == CellType.BLANK) { // nếu trống -> thêm 0.0
                            workingDaysByDate.add(0.0);
                        }
                    }
                }
            }

        }
        return workingDaysByDate;
    }

    //     tìm vị trí cột dòng chứa giá trị value( số hoặc chuỗi)
    public int findPositionCell(Sheet sheet, Integer value) {
        int indexColumn = -1;
        for (Row row : sheet) {
            if(value instanceof Integer) { // nếu là số kiểm tra kiểu numeric
                if(indexColumn==-1){
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.NUMERIC && (int)cell.getNumericCellValue() ==(int) value) {
                            indexColumn = cell.getColumnIndex();
                            break;
                        }
                    }
                    continue;
                }
            }
//            else if(value instanceof String) { // nếu là chuỗi -> kiểm tra kiểu String
//                if(indexColumn==-1){
//                    for (Cell cell : row) {
//                        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(value.toString())) {
//                            indexColumn = cell.getRowIndex();
//                        }
//                    }
//                    continue;
//                }
//            }

        }
        return indexColumn;
    }

    public int findPositionCell(Sheet sheet, String value) {
        int indexColumn = -1;
        for (Row row : sheet) {
            if(value instanceof String) { // nếu là chuỗi -> kiểm tra kiểu String
                if(indexColumn==-1){
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(value.toString())) {
                            indexColumn = cell.getRowIndex();
                        }
                    }
                    continue;
                }
            }

        }
        return indexColumn;
    }



    //     kiểm tra ô tại dòng 5, cột position có chứa chuỗi bắt đầu bằng findInfo
    public boolean checkDaySunDay(Sheet sheet,int position,String findInfo) {
        Cell cellTypeHourWork = sheet.getRow(5).getCell(position, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        if (cellTypeHourWork.getCellType() == CellType.STRING && cellTypeHourWork.getStringCellValue().startsWith(findInfo) ) {
            return true;
        }
        return false;
    }

    // lấy giá trị cột Q
    public List<Double> extractColumnQ(Sheet sheet, int startRow, int endRow, Workbook workbook) {
        List<Double> columnQValues = new ArrayList<>();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        for (int i = startRow; i < endRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                columnQValues.add(0.0);
                continue;
            }
            Cell cell = row.getCell(16, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            double value = getCellNumeric(cell, evaluator);
            columnQValues.add(value);
        }

        return columnQValues;
    }

    //  xử lý các loại ô -> trả về giá trị số. nếu là ô số -> sử dụng FormulaEvaluator để tính giá trị
    private double getCellNumeric(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return 0.0;
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (cell.getCellType() == CellType.FORMULA) {
            try {
                return evaluator.evaluate(cell).getNumberValue();
            } catch (Exception e) {
                return 0.0;
            }
        }
        return 0.0;
    }
}