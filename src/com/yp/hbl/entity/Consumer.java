package com.yp.hbl.entity;

public class Consumer {
    //店别
    private String shopNo;
    //区域
    private String area;
    //客户编号
    private String consumerNum;
    //客户名称
    private String consumerName;
    //订单号码
    private String orderNum;
    //收货日期
    private String deliveryDate;
    //客户未税金额
    private double consumerNoTaxAmount;
    //折扣
    public static final double DISCOUNT = 0.1287;

    public Consumer() {
    }

    public Consumer(String area, String consumerNum, String consumerName, String orderNum, String deliveryDate, double consumerNoTaxAmount) {
        this.area = area;
        this.consumerNum = consumerNum;
        this.consumerName = consumerName;
        this.orderNum = orderNum;
        this.deliveryDate = deliveryDate;
        this.consumerNoTaxAmount = consumerNoTaxAmount;
    }

    public Consumer(String consumerName, String orderNum) {
        this.consumerName = consumerName;
        this.orderNum = orderNum;
    }


    public String getShopNo() {
        return shopNo;
    }

    public void setShopNo(String shopNo) {
        this.shopNo = shopNo;
    }

    public String getArea() {
        return area;
    }

    public void setArea(String area) {
        this.area = area;
    }

    public String getConsumerNum() {
        return consumerNum;
    }

    public void setConsumerNum(String consumerNum) {
        this.consumerNum = consumerNum;
    }

    public String getConsumerName() {
        return consumerName;
    }

    public void setConsumerName(String consumerName) {
        this.consumerName = consumerName;
    }

    public String getOrderNum() {
        return orderNum;
    }

    public void setOrderNum(String orderNum) {
        this.orderNum = orderNum;
    }

    public String getDeliveryDate() {
        return deliveryDate;
    }

    public void setDeliveryDate(String deliveryDate) {
        this.deliveryDate = deliveryDate;
    }

    public double getConsumerNoTaxAmount() {
        return consumerNoTaxAmount;
    }

    public void setConsumerNoTaxAmount(double consumerNoTaxAmount) {
        this.consumerNoTaxAmount = consumerNoTaxAmount;
    }

    @Override
    public String toString() {
        return "Consumer{" +
                "area='" + area + '\'' +
                ", consumerNum='" + consumerNum + '\'' +
                ", consumerName='" + consumerName + '\'' +
                ", orderNum='" + orderNum + '\'' +
                ", deliveryDate='" + deliveryDate + '\'' +
                ", consumerNoTaxAmount=" + consumerNoTaxAmount +
                '}';
    }
}
