package software.chronicle.xlsx2java;

import static software.chronicle.xlsx2java.ExcelMath.*;

/**
 * Created by peter on 19/07/2017.
 */
public class XLSXSample {
    double B3, C3;
    double B4;
    double B5;
    double B6;
    double B7;
    double B8;
    double B9;

    public void reset() {
        B3 = 1.0;
        C3 = 1.5;
        B4 = Double.NaN;
        B5 = Double.NaN;
        B6 = Double.NaN;
        B7 = Double.NaN;
        B8 = Double.NaN;
        B9 = Double.NaN;
    }

    public void onePass() {
        B4 = MIN(B3, C3);
        B5 = MAX(B3, C3);
        B6 = IF(B3 > C3, B3, C3);
        B7 = ROUND(C3, 0);
        B8 = CEILING(C3, 0.1);
        B9 = FLOOR(C3, 0.1);
    }

    public void twoPass() {
    }

    public String toCsv() {
        return "," + B3 + "," + C3 + "\n" +
                "," + B4 + "\n" +
                "," + B5 + "\n" +
                "," + B6 + "\n" +
                "," + B7 + "\n" +
                "," + B8 + "\n" +
                "," + B9 + "\n";
    }

}
