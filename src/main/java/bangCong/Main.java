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

            // Lấy giá trị cột Q
            List<Double> columnQValues = excelService.extractColumnQ(sheet, 6, 6 + listNameNVs.size(), workbook);

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
            for (int i = 0; i < employeeInfo.size(); i++) {
                Employee emp = employeeInfo.get(i);
                double colQ = columnQValues.get(i);
                double difference = Math.abs(emp.getTotalPrice() - colQ);
                String comparison = (difference < 0.01) ? "Khớp" : "Không khớp (chênh lệch: " + dc.format(difference) + ")";
                System.out.println(
                        "Nhân viên:\n" +
                                "\tMã NV: " + emp.getId() + "\n" +
                                "\tTên: " + emp.getName() + "\n" +
                                "\tTổng số giờ làm: " + emp.getTotalHoursWorked() + "\n" +
                                "\tTổng tiền các ca làm: " + dc.format(emp.getTotalPrice()) + "\n" +
                                "\tTổng tiền cột Q: " + dc.format(colQ) + "\n" +
                                "\tSo sánh với cột Q: " + comparison + "\n"
                );
            }

            System.out.print("Nhập ngày đầu để tìm kiếm: ");
            Scanner sc = new Scanner(System.in);
            int startDay = sc.nextInt();
            System.out.print("Nhập ngày cuối để tìm kiếm: ");
            int endDay = sc.nextInt();
            System.out.println("Vị trí ngày đầu tìm thấy ở cột thứ: "+ excelService.findPositionCell(sheet,startDay));
            System.out.println("Vị trí ngày cuối tìm thấy ở cột thứ: "+(excelService.findPositionCell(sheet,endDay+1)-1));
            List<Double>[] arrayListHourseWork = new List[listNameNVs.size()];

            for(int i=0;i<listNameNVs.size();i++){
                List<Double> countWorkingDaysByDate = new ArrayList<>();
                double totalHourseWork = 0;
//                int findPositionCellString = excelService.findPositionCell(sheet,listNameNVs.get(i));
                int findPositionCellString = excelService.findPositionCell(sheet, listNameNVs.get(i));
                countWorkingDaysByDate.addAll(excelService.countWorkingDaysByDate(sheet, excelService.findPositionCell(sheet,startDay),findPositionCellString, excelService.findPositionCell(sheet,endDay+1)-1));
                arrayListHourseWork[i] = countWorkingDaysByDate;
                Long countWorkEmployee = countWorkingDaysByDate.stream().filter(x -> x>0).count();
                for(int j=0;j<countWorkingDaysByDate.size();j++){
                    totalHourseWork += countWorkingDaysByDate.get(j);
                }
                //ngày có làm việc - lương mỗi ngày
                System.out.println(
                        "Nhân viên:\n" +
                                "\tTên: " + listNameNVs.get(i) + "\n" +
                                "\tMã NV: " + listMaNVs.get(i) + "\n" +
                                "\tThời gian làm việc: được tìm kiếm từ ngày " + startDay + " đến hết ngày " + endDay + "\n" +
                                "\tTổng số ngày làm: " + countWorkEmployee + " ngày\n" +
                                "\tTổng số giờ làm: " + totalHourseWork + " giờ\n"
                );
            }

            // luu cac ngay chu nhat vi tri cot
            List<Integer> daySundays = new ArrayList<>();
            // luu cac thu la ngay chu nhat (1->5) thì luu 2 la ngay cn
            for(int m = startDay;m<=endDay;m++){
                int findPosition = excelService.findPositionCell(sheet,m);
                if(excelService.checkDaySunDay(sheet,findPosition,"WK")){
                    daySundays.add(m);
                }
            }

//            System.out.println("Tìm kiếm từ ngày "+startDay+" đến hết ngày "+endDay+" có tổng số ngày chủ nhật: "+daySundays.size());
            for (int j = 0; j < listNameNVs.size(); j++) {
                System.out.println(String.format("\n%-2s - %-1s", "Nhân viên", "Mã NV"));
                Map<String,Double> infoCa = map.get(listMaNVs.get(j));
                System.out.println(String.format("%-2s - %-1s", listNameNVs.get(j), listMaNVs.get(j)));
                System.out.println("-".repeat(20));
                int k = 0;
                for (int i = startDay; i < endDay; i++) {
                    double sumPriceInDay = 0;
                    double totalHoursWork = 0;
                    if (daySundays.contains(i)) {
                        Double[] x = new Double[tongCaDaySunDay];
                        for (int n = 0; n < tongCaDaySunDay; n++) {
//                            System.out.println("check arrayListHourseWork[j]  " + arrayListHourseWork[j]);
                            if(k<arrayListHourseWork[j].size()) {
                                x[n] = arrayListHourseWork[j].get(k);
                            }else{
                                x[n] = 0.0;
                            }
                            totalHoursWork += x[n];
                            if (x[n] > 0) {
                                Double tienLamTrongCa = infoCa.get(listMaNVs.get(j)+listCaDaySunday.get(n))*x[n];
                                if(infoCa.get(listMaNVs.get(j)+listCaDaySunday.get(n))==null) {
                                    System.out.print("ERROR");
                                }
                                System.out.println(String.format("- Ngày thứ %-2d - Ca: %-10s - Số giờ: %.2f - Tổng tiền: %.2f",
                                        i, listCaDaySunday.get(n), x[n],tienLamTrongCa));
                                sumPriceInDay += tienLamTrongCa;
                            }
                            k += 1;
                        }
                    } else {
                        Double[] x = new Double[tenCacCa.size()];
                        for (int m = 0; m < tongCaNgayThuong; m++) {
                            if(k<arrayListHourseWork[j].size()) {
                                x[m] = arrayListHourseWork[j].get(k);
                            }else{
                                x[m] = 0.0;
                            }

                            totalHoursWork += x[m];
                            if (x[m] > 0) {
                                Double tienLamTrongCa = infoCa.get(listMaNVs.get(j)+listCaNgayThuong.get(m))*x[m];
                                System.out.println(String.format("-Ngày thứ %-3d - Ca: %-1s - Số giờ: %.2f - Tổng tiền: %.2f",
                                        i, listCaNgayThuong.get(m), x[m],tienLamTrongCa));
                                sumPriceInDay+= tienLamTrongCa;
                            }
                            k += 1;
                        }
                    }
                    System.out.println(String.format("Tổng số giờ làm trong ngày thứ %-3d: %.2f", i, totalHoursWork));
                    System.out.println(String.format("Tổng tien làm trong ngày thứ %-3d: %.2f", i, sumPriceInDay));
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}