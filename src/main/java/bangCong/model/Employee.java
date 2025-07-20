package bangCong.model;

import java.util.ArrayList;
import java.util.List;

public class Employee {
    private String id;
    private String name;
    private Double totalHoursWorked;
    private Integer totalDaysWorked;
    private Double totalEaring;

    public Employee() {
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
}
