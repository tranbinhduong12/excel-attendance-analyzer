package bangCong;

import bangCong.model.Employee;
import bangCong.service.ExcelService;
import bangCong.service.ExcelServiceImpl;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.*;

public class Main {
    public static void main(String[] args) throws Exception {
        List<Double> giaCacCa = new ArrayList<>();
        ExcelService excelService = new ExcelServiceImpl();
        List<Employee> employeeInfo = new ArrayList<>();
        List<String> listNameNVs = new ArrayList<>();
        List<String> listMaNVs = new ArrayList<>();
        List<Double> listHourseCas = new ArrayList<>();
        List<Double> listPriceCas = new ArrayList<>();
//        List<String> allowedExtensions = Arrays.asList("xlsx", "xls");
        List<String> tenCacCa = new ArrayList<>();
        String excelFilePath = "./BangCong.xlsx";
        File file = new File(excelFilePath);
        System.out.println(file.getName());
//        if(!FileValidator.isValidFile(file,allowedExtensions)) {
//            throw new Exception("File không hợp lệ");
//        }
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // đọc sheet đầu tiên
            // goi ham lay danh sach ten nhan vien
            listNameNVs = excelService.extractEmployeeInfo(sheet, "Họ Tên", listNameNVs);
            // goi ham lay danh sach ma nhan vien
            listMaNVs = excelService.extractEmployeeInfo(sheet, "Mã NV", listMaNVs);
            // goi ham lay danh sach cac ca truc
            List<String> listCaNgayThuong = excelService.listCaNgayThuong(sheet, 3, 15);
            List<String> listCaDaySunday = excelService.listCaDaySunday(sheet, 3, 15);
            int tongCaDaySunDay = excelService.listCaDaySunday(sheet, 3, 15).size();
            int tongCaNgayThuong = listCaNgayThuong.size();
            tenCacCa.addAll(listCaNgayThuong);
            tenCacCa.addAll(listCaDaySunday);

            // lấy giá trị thời gian các ca
            for (int i = 0; i < tenCacCa.size(); i++) {
                List<Double> listHourseCa = new ArrayList<>();
                listHourseCa = excelService.totalHourseWork(sheet, 6, tenCacCa.get(i).toString(), 6 + listNameNVs.size());
                listHourseCas.addAll(listHourseCa);
            }
            // lấy giá trị tiền các ca
            for (int i = 0; i < tenCacCa.size(); i++) {
                List<Double> listPriceCa = new ArrayList<>();
                listPriceCa = excelService.totalPriceCa(sheet, 6, tenCacCa.get(i).toString(), 6 + listNameNVs.size());
                listPriceCas.addAll(listPriceCa);
            }
            // chưa hiểu
            Map<String, Map<String, Double>> map = new HashMap<>();
            for (int i = 0; i < listNameNVs.size(); i++) { // in ra ds nvien
                System.out.println("NhanVien: " + listNameNVs.get(i));
                for (int caIndex = 0; caIndex < tenCacCa.size(); caIndex++) { // k hiểu
                    int viTri = caIndex * listNameNVs.size() + i; // Tính đúng vị trí giá trong listPriceCas
                    if (viTri < listPriceCas.size()) {//
                        String tenCa = tenCacCa.get(caIndex); // lấy tên ca theo thứ tự
                        Double giaCa = listPriceCas.get(viTri); // lấy giá ứng với nhân viên và ca đó
                        System.out.println(String.format("  TenCa: %-1s - GiaCa: %.2f", tenCa, giaCa));
                        Map<String, Double> infoCa = map.getOrDefault(listMaNVs.get(i), new LinkedHashMap<>());// chưa hiểu
                        infoCa.put(listMaNVs.get(i) + tenCa, giaCa);
                        map.put(listMaNVs.get(i), infoCa);
                    }
                }

            }
            //tính tổng số giờ làm mỗi nv
            for (int i = 0; i < listNameNVs.size(); i++) {
                Employee infoEmployee = new Employee();
                infoEmployee.setId(listMaNVs.get(i));
                infoEmployee.setName(listNameNVs.get(i));
                double sum = 0, price = 0; // tổng  giờ làm, tổng giá
                for (int j = 0; j < listHourseCas.size(); j++) {
                    if (j % (listNameNVs.size()) == i) {
                        sum = sum + listHourseCas.get(j);
                        price = listHourseCas.get(j) * listPriceCas.get(j) + price;
                        giaCacCa.add(listPriceCas.get(j));
                    }
                }
                infoEmployee.setTotalPrice(price);
                infoEmployee.setTotalHoursWorked(sum);
                infoEmployee.setPriceCacCa(giaCacCa);
                employeeInfo.add(infoEmployee);
            }
            System.out.print("Tổng số giờ làm việc của các nhân viên là: ");
            System.out.println();
            DecimalFormat dc = new DecimalFormat("0.00");
            employeeInfo.stream().forEach(x -> System.out.println(
                    "Nhân viên:\n" +
                            "\tTên: " + x.getName() + "\n" +
                            "\tMã NV: " + x.getId() + "\n" +
                            "\tTổng số giờ làm: " + x.getTotalHoursWorked() + "\n" +
                            "\tTổng tiền các ca làm: " + dc.format(x.getTotalPrice()) + "\n"
            ));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}