package com.yp.hbl.entity;

public class FkData {
    //店别
    private int shopNo;
    //店编号
    private String shopNum;
    //店名
    private String shopName;
    //产品编号
    private String productNo;
    //产品名称
    private String productName;
    //公司单价=公司数据分项 单价
    private double companyPrice;
    //客户单价=公司数据分项 对方价 最后日期
    private double consumerPrice;
    //产品数量
    private int quantity;
    //差异金额
    private double balanceAmount;
    //菜单内差异金额
    private double innerBalanceAmount;
    //菜单外差异金额
    private double outerBalanceAmount;
    //时间段
    private String dateRange;

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

    public String getProductNo() {
        return productNo;
    }

    public void setProductNo(String productNo) {
        this.productNo = productNo;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public double getCompanyPrice() {
        return companyPrice;
    }

    public void setCompanyPrice(double companyPrice) {
        this.companyPrice = companyPrice;
    }

    public double getConsumerPrice() {
        return consumerPrice;
    }

    public void setConsumerPrice(double consumerPrice) {
        this.consumerPrice = consumerPrice;
    }

    public int getQuantity() {
        return quantity;
    }

    public void setQuantity(int quantity) {
        this.quantity = quantity;
    }

    public double getBalanceAmount() {
        return balanceAmount;
    }

    public void setBalanceAmount(double balanceAmount) {
        this.balanceAmount = balanceAmount;
    }

    public double getInnerBalanceAmount() {
        return innerBalanceAmount;
    }

    public void setInnerBalanceAmount(double innerBalanceAmount) {
        this.innerBalanceAmount = innerBalanceAmount;
    }

    public double getOuterBalanceAmount() {
        return outerBalanceAmount;
    }

    public void setOuterBalanceAmount(double outerBalanceAmount) {
        this.outerBalanceAmount = outerBalanceAmount;
    }

    public String getDateRange() {
        return dateRange;
    }

    public void setDateRange(String dateRange) {
        this.dateRange = dateRange;
    }

    @Override
    public String toString() {
        return "FkData{" +
                "shopNo=" + shopNo +
                ", shopNum='" + shopNum + '\'' +
                ", shopName='" + shopName + '\'' +
                ", productNo='" + productNo + '\'' +
                ", productName='" + productName + '\'' +
                ", companyPrice=" + companyPrice +
                ", consumerPrice=" + consumerPrice +
                ", quantity=" + quantity +
                ", balanceAmount=" + balanceAmount +
                ", innerBalanceAmount=" + innerBalanceAmount +
                ", outerBalanceAmount=" + outerBalanceAmount +
                ", dateRange='" + dateRange + '\'' +
                '}';
    }
}
