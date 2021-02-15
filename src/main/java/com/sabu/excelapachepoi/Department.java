package com.sabu.excelapachepoi;


import lombok.Data;
import lombok.Getter;
import lombok.Setter;

@Data
@Getter
@Setter
public class Department {

    private String name;

    private Integer opdNo;

    private String startTime;

    private String endTime;

    private String days;
}
