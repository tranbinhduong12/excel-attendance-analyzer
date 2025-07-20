package bangCong.service;

import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.IOException;

public interface ExcelService {
    // đọc file
    Sheet readExcelFile(File file) throws IOException;
}
