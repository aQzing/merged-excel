package com.qzing.po;

import java.util.List;
import java.util.Map;

public class ReportPO {
    //报表名
    private String name;
    //维度 维度-数据库表名
    private Map dimension;


    public void setName(String name) {
        this.name = name;
    }

    public void setDimension(Map dimension) {
        this.dimension = dimension;
    }

    public String getName() {
        return name;
    }

    public Map getDimension() {
        return dimension;
    }
}
