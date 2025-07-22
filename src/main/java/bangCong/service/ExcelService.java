package bangCong.service;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

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

}