package bangCong.model;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Employee {
    private String id;
    private String name;
    private Double totalHoursWorked;
    private Integer totalDaysWorked;
    private Double totalPrice;
    private List<Double> priceCacCa;

//    private List<Double> priceMoiCa;

    public Employee() {

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

    public Double getTotalPrice() {
        return totalPrice;
    }

    public void setTotalPrice(Double totalPrice) {
        this.totalPrice = totalPrice;
    }

    public List<Double> getPriceCacCa() {
        return priceCacCa;
    }

    public void setPriceCacCa(List<Double> priceCacCa) {
        this.priceCacCa = priceCacCa;
    }
}