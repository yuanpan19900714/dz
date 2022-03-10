package com.yp.hbl.entity;

public class Dz {
    //店名编号
    private String shopNameNum;
    //店名
    private String shopName;
    //单号
    private String orderNum;
    //日期
    private String date;
    //产品编号
    private String productNum;
    //产品名称
    private String productName;
    //单价
    private double untiPrice;
    //数量
    private int quantity;
    //金额
    private double amount;
    //对方价
    private double otherPrice;
    //固定倍数
    public static final double BS = 1.13;
    //对方单价
    private double otherUnitPrice;
    //差价
    private double disparity;
    //差额
    private double balance;

    public Dz() {
    }

    public Dz(String shopNameNum, String date, String productNum) {
        this.shopNameNum = shopNameNum;
        this.date = date;
        this.productNum = productNum;
    }

    @Override
    public String toString() {
        return "Dz{" +
                "shopNameNum='" + shopNameNum + '\'' +
                ", shopName='" + shopName + '\'' +
                ", orderNum='" + orderNum + '\'' +
                ", date='" + date + '\'' +
                ", productNum='" + productNum + '\'' +
                ", productName='" + productName + '\'' +
                ", untiPrice=" + untiPrice +
                ", quantity=" + quantity +
                ", amount=" + amount +
                ", otherPrice=" + otherPrice +
                ", BS=" + BS +
                ", otherUnitPrice=" + otherUnitPrice +
                ", disparity=" + disparity +
                ", balance=" + balance +
                '}';
    }

    public Dz(String shopNameNum, String shopName, String orderNum, String date, String productNum, String productName, double untiPrice, int quantity, double amount, double otherPrice, double otherUnitPrice, double disparity, double balance) {
        this.shopNameNum = shopNameNum;
        this.shopName = shopName;
        this.orderNum = orderNum;
        this.date = date;
        this.productNum = productNum;
        this.productName = productName;
        this.untiPrice = untiPrice;
        this.quantity = quantity;
        this.amount = amount;
        this.otherPrice = otherPrice;
        this.otherUnitPrice = otherUnitPrice;
        this.disparity = disparity;
        this.balance = balance;
    }

    public String getShopNameNum() {
        return shopNameNum;
    }

    public void setShopNameNum(String shopNameNum) {
        this.shopNameNum = shopNameNum;
    }

    public String getShopName() {
        shopName = getString(shopName);
        return shopName;
    }

    public void setShopName(String shopName) {
        this.shopName = shopName;
    }

    public String getOrderNum() {
        return orderNum;
    }

    public void setOrderNum(String orderNum) {
        this.orderNum = orderNum;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getProductNum() {
        return productNum;
    }

    public void setProductNum(String productNum) {
        this.productNum = productNum;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public double getUntiPrice() {
        return untiPrice;
    }

    public void setUntiPrice(double untiPrice) {
        this.untiPrice = untiPrice;
    }

    public int getQuantity() {
        return quantity;
    }

    public void setQuantity(int quantity) {
        this.quantity = quantity;
    }

    public double getAmount() {
        return amount;
    }

    public void setAmount(double amount) {
        this.amount = amount;
    }

    public double getOtherPrice() {
        return otherPrice;
    }

    public void setOtherPrice(double otherPrice) {
        this.otherPrice = otherPrice;
    }

    public double getBS() {
        return BS;
    }

    public double getOtherUnitPrice() {
        return otherUnitPrice;
    }

    public void setOtherUnitPrice(double otherUnitPrice) {
        this.otherUnitPrice = otherUnitPrice;
    }

    public double getDisparity() {
        return disparity;
    }

    public void setDisparity(double disparity) {
        this.disparity = disparity;
    }

    public double getBalance() {
        return balance;
    }

    public void setBalance(double balance) {
        this.balance = balance;
    }

    public static String dealShopName(String shopName) {
        shopName = getString(shopName);
        return shopName;
    }

    private static String getString(String shopName) {
        if (shopName.contains("---")) {
            shopName = shopName.replace("---", "-");
        } else if (shopName.contains("--")) {
            shopName = shopName.replace("--", "-");
        } else if (shopName.contains("—")) {
            shopName = shopName.replace("—", "-");
        }
        if (shopName.contains("杭州")) {
            shopName = shopName.replace("杭州", "");
        }
        if (!shopName.contains("店")) {
            System.out.println("shopName = [" + shopName + "]");
        }
        if (shopName.contains("店") && !shopName.endsWith("店")) {
            shopName = shopName.substring(0, shopName.lastIndexOf("店") + 1);
        }
        return shopName;
    }
}
