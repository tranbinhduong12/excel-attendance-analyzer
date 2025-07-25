package bangCong.service;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


import java.util.List;
import java.util.Map;

public interface ExcelService {

    // lấy danh sách ca ngày thường
    public List<String> listCaNgayThuong(Sheet sheet, int startIxColumn, int endIxColumn);

    // lấy danh sách ca chủ nhật
    public List<String> listCaDaySunday(Sheet sheet,int startIxColumn,int endIxColumn);

    // lấy thông tin nv
    public List<String> extractEmployeeInfo(Sheet sheet,String findInfo,List<String> infoNV);

    // tổng giờ làm mỗi ca
    public List<Double> totalHourseWork(Sheet sheet,int indexStarRow,String findInfo,int indexEndRow);

    // tổng số tiền mỗi ca
    public List<Double> totalPriceCa(Sheet sheet,int indexStarRow,String findInfo,int indexEndRow);

    // lấy giá trị cột Q
    public List<Double> extractColumnQ(Sheet sheet, int startRow, int endRow, Workbook workbook);

    // lấy số giờ làm của nhân viên trong khoảng ngày được chọn
    public List<Double> countWorkingDaysByDate(Sheet sheet,int indexStartColumn,int indexRow,int indexEndColumn);

    //tìm vị trí của cột ngày bắt đầu và ngày kết thúc trong sheet
    public int findPositionCell(Sheet sheet, Integer value);


    public int findPositionCell(Sheet sheet, String value);

    //     kiểm tra ngày chủ nhật
    public boolean checkDaySunDay(Sheet sheet,int position,String findInfo);

}