package bangCong.service;

import bangCong.model.Employee;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.IOException;
import java.util.List;

public interface ExcelService {
    // đọc file
    Sheet readExcelFile(File file) throws IOException;

    // lấy danh sách thông tin nhân viên
    public List<String> employeeInfo(Sheet sheet,String findInfo,List<String> infoEmployee);
}
