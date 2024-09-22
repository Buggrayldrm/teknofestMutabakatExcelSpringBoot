package org.example.excelproject.entities;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Data;

@Entity
@Data
public class DataEntity {
    @Id
    private Long id;
    private String tc;
    private String city;
    private String date;

    private int lunchSubcontractor;
    private int lunchStaff;
    private int dinnerSubcontractor;
    private int dinnerStaff;
    private int nightFood;
    private int lunchbox;


    private double lunchSubcontractorPrice;
    private double lunchStaffPrice;
    private double dinnerSubcontractorPrice;
    private double dinnerStaffPrice;
    private double nightFoodPrice;
    private double lunchboxPrice;

}
