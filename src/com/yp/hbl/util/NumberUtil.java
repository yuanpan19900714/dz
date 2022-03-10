package com.yp.hbl.util;

import java.math.BigDecimal;
import java.math.RoundingMode;

public class NumberUtil {
    public static double getRound(double db, int scale) {
        BigDecimal bd = new BigDecimal(db);
        return bd.setScale(scale, RoundingMode.HALF_UP).doubleValue();
    }
}
