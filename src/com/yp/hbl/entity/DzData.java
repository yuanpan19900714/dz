package com.yp.hbl.entity;

public class DzData {
    //店别
    private int shopNo;
    //店编号
    private String shopNum;
    //店名
    private String shopName;
    //日期
    private String date;
    //订单号码
    private String orderNum;
    //公司金额=sum(公司数据分项：店名—日期—订单号码)
    private double companyAmount;
    //客户金额=客户数据：客户未税金额
    private double consumerAmount;
    //客户含税金额=客户金额*1.13
    private double consumerTaxAmount;
    //折扣金额=客户含税金额*0.1287
    private double discountAmount;
    //开票金额=客户含税金额-折扣金额
    private double billAmount;
    //差异金额=公司金额-客户含税金额
    private double balanceAmount;
    //客户数据和公司数据是否对应
    private boolean isSame;

    public DzData() {
    }

    public int getShopNo() {
        return shopNo;
    }

    public void setShopNo(int shopNo) {
        this.shopNo = shopNo;
    }

    public String getShopNum() {
        return shopNum;
    }

    public void setShopNum(String shopNum) {
        this.shopNum = shopNum;
    }

    public String getShopName() {
        return shopName;
    }

    public void setShopName(String shopName) {
        this.shopName = shopName;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getOrderNum() {
        return orderNum;
    }

    public void setOrderNum(String orderNum) {
        this.orderNum = orderNum;
    }

    public double getCompanyAmount() {
        return companyAmount;
    }

    public void setCompanyAmount(double companyAmount) {
        this.companyAmount = companyAmount;
    }

    public double getConsumerAmount() {
        return consumerAmount;
    }

    public void setConsumerAmount(double consumerAmount) {
        this.consumerAmount = consumerAmount;
    }

    public double getConsumerTaxAmount() {
        return consumerTaxAmount;
    }

    public void setConsumerTaxAmount(double consumerTaxAmount) {
        this.consumerTaxAmount = consumerTaxAmount;
    }

    public double getDiscountAmount() {
        return discountAmount;
    }

    public void setDiscountAmount(double discountAmount) {
        this.discountAmount = discountAmount;
    }

    public double getBillAmount() {
        return billAmount;
    }

    public void setBillAmount(double billAmount) {
        this.billAmount = billAmount;
    }

    public double getBalanceAmount() {
        return balanceAmount;
    }

    public void setBalanceAmount(double balanceAmount) {
        this.balanceAmount = balanceAmount;
    }

    public boolean isSame() {
        return isSame;
    }

    public void setSame(boolean same) {
        isSame = same;
    }

    @Override
    public String toString() {
        return "DzData{" +
                "shopNo='" + shopNo + '\'' +
                ", shopNum='" + shopNum + '\'' +
                ", shopName='" + shopName + '\'' +
                ", date='" + date + '\'' +
                ", orderNum='" + orderNum + '\'' +
                ", companyAmount=" + companyAmount +
                ", consumerAmount=" + consumerAmount +
                ", consumerTaxAmount=" + consumerTaxAmount +
                ", discountAmount=" + discountAmount +
                ", billAmount=" + billAmount +
                ", balanceAmount=" + balanceAmount +
                '}';
    }
}
