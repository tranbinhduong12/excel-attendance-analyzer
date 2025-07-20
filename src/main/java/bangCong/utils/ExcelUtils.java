package bangCong.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

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

}
