package com.yp.hbl.entity;

public class Constant {
    public class dataCompareWorkbook {
        public final static String WORKBOOK_NAME = "数据对照表";
        public final static String HIS_SHEET_NAME = "上期数据";
        public final static String SRC_SHEET_NAME = "公司数据分项";
        public final static String CONSUMER_SHEET_NAME = "客户数据";
    }

    public class dzWorkbook {
        public final static String WORKBOOK_NAME = "对账工作簿";
        public final static String SUM_SHEET_NAME = "汇总一览表";
        public final static int MAX_SHEET_NUM = 40;
    }

    public class fkWorkbook {
        public final static String WORKBOOK_NAME = "反馈工作簿";
        public final static String SUM_SHEET_NAME = "导航页";
        public final static int MAX_SHEET_NUM = 40;
    }

    public enum ShopNo {
        JX("DRF-嘉兴店", 6),
        HZ("DRF-湖州店", 15),
        XS("DRF-萧山店", 26),
        FY("DRF-富阳店", 29),
        SY("DRF-上虞店", 41),
        ZJ("DRF-诸暨店", 43),
        PH("DRF-平湖店", 50),
        TL("DRF-桐庐店", 57),
        HY("DRF-海盐店", 61),
        JS("DRF-嘉善店", 77),
        SQ("DRF-石桥店", 80),
        JH("DRF-金华店", 116),
        SX("DRF-市心(润福)店", 130),
        PJ("DRF-袍江店", 142),
        KQ("DRF-柯桥店", 152),
        ZH("DRF-中环西路店", 165),
        YH("DRF-余杭店", 189),
        JSH("DRF-泾水店", 712),
        DG("DRF-大关店", 721),
        SL("DRF-胜利店", 722),
        CX("DRF-长兴店", 723),
        SL1("DRF-胜利1店", 1),
        FY1("DRF-富阳1店", 21),
        FY2("DRF-富阳2店", 13),
        JX3("DRF-嘉兴3店", 14),
        SQ1("DRF-石桥1店", 15),
        TL1("DRF-桐庐1店", 16),
        TL3("DRF-桐庐3店", 17),
        YH1("DRF-余杭1店", 18),
        TL2("DRF-桐庐2店", 19),
        JX1("DRF-嘉兴1店", 10),
        JS1("DRF-嘉善1店", 101),
        HZ1("DRF-湖州1店", 102),
        HZ5("DRF-湖州5店", 103),
        HZ2("DRF-湖州2店", 104),
        JX2("DRF-嘉兴2店", 105),
        HY1("DRF-海盐1店", 107),
        YH2("DRF-余杭2店", 108),
        HZ3("DRF-湖州3店", 109),
        HZ4("DRF-湖州4店", 1110),
        HZ6("DRF-湖州6店", 1111),
        JH1("DRF-金华1店", 1112),
        ZJ1("DRF-诸暨1店", 1113),
        HZ7("DRF-湖州7店", 1114),
        JS2("DRF-嘉善2店", 1115),
        ZJ2("DRF-诸暨2店", 1116);

        private String shopName;
        private int shopNo;

        ShopNo(String shopName, int shopNo) {
            this.shopName = shopName;
            this.shopNo = shopNo;
        }


        public String getShopName() {
            return shopName;
        }

        public void setShopName(String shopName) {
            this.shopName = shopName;
        }

        public int getShopNo() {
            return shopNo;
        }

        public void setShopNo(int shopNo) {
            this.shopNo = shopNo;
        }

        public static int getShopNo(String shopName) {
            for (ShopNo sn : ShopNo.values()) {
                if (sn.getShopName().equals(shopName)) {
                    return sn.getShopNo();
                }
            }
            return 0;
        }
    }
}
