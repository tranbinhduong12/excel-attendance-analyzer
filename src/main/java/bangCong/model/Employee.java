package bangCong.model;

import java.util.List;

public class Employee {
    private String id;
    private String name;
    private Double totalHoursWorked; // tổng giờ làm = GC + TC + GC1 + TC1 + WK-D + WK-N
    private Integer totalDaysWorked; // tổng lương = (GC*giáGC + TC*giáTC + ...) + tiền WK
    private Double totalEaring;
    private List<Double> priceMoiCa;

    public Employee(String id, String name) {
        this.id = id;
        this.name = name;
    }

    public Employee(String id, String name, Double totalHoursWorked, Integer totalDaysWorked, Double totalEaring) {
        this.id = id;
        this.name = name;
        this.totalHoursWorked = totalHoursWorked;
        this.totalDaysWorked = totalDaysWorked;
        this.totalEaring = totalEaring;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getTotalHoursWorked() {
        return totalHoursWorked;
    }

    public void setTotalHoursWorked(Double totalHoursWorked) {
        this.totalHoursWorked = totalHoursWorked;
    }

    public Integer getTotalDaysWorked() {
        return totalDaysWorked;
    }

    public void setTotalDaysWorked(Integer totalDaysWorked) {
        this.totalDaysWorked = totalDaysWorked;
    }

    public Double getTotalEaring() {
        return totalEaring;
    }

    public void setTotalEaring(Double totalEaring) {
        this.totalEaring = totalEaring;
    }

    private Double actualTotalFromExcel; // so sánh tổng tiền tính đc với tổng tiền ở cột Q

    public void setActualTotalFromExcel(Double actualTotalFromExcel) {
        this.actualTotalFromExcel = actualTotalFromExcel;
    }

    public Double getActualTotalFromExcel() {
        return actualTotalFromExcel;
    }

//    @Override
//    public String toString() {
//        return "employee: " + id + " - " + name + "\n"
//                + " Tổng giờ: " + totalHoursWorked + "h\n"
//                + " Tổng tiền: " + String.format("%,.0f", totalEaring) + " VND\n";
//    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("Employee: ").append(id).append(" - ").append(name).append("\n");
        sb.append(" Tổng giờ: ").append(totalHoursWorked).append("h\n");
        sb.append(" Tổng tiền (tính): ").append(String.format("%,.0f", totalEaring)).append(" VND\n");
        sb.append(" Tổng tiền (trong file): ").append(String.format("%,.0f", actualTotalFromExcel)).append(" VND\n");

        if (Math.abs(totalEaring - actualTotalFromExcel) > 1e-2) {
            sb.append("Kết quả: Không khớp\n");
        } else {
            sb.append("Kết quả: Khớp\n");
        }

        return sb.toString();
    }
}