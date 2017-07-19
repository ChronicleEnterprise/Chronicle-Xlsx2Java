package software.chronicle.xlsx2java;

import net.openhft.chronicle.core.Maths;

import java.util.Arrays;
import java.util.concurrent.ThreadLocalRandom;

/**
 * Created by peter on 19/07/2017.
 */
public interface ExcelMath {

    public static double ABS(double a) {
        return Math.abs(a);
    }

    public static double MAX(double a, double b) {
        return Math.max(a, b);
    }

    public static double MIN(double a, double b) {
        return Math.min(a, b);
    }

    public static double IF(boolean test, double a, double b) {
        return test ? a : b;
    }

    public static double FLOOR(double a, double precision) {
        double factor = Math.round(1 / precision);
        return Math.floor(a * factor) / factor;
    }

    public static double CEILING(double a, double precision) {
        double factor = Math.round(1 / precision);
        return Math.ceil(a * factor) / factor;
    }

    public static double ROUND(double a, int precision) {
        return Maths.roundN(a, precision);
    }

    public static double SIGN(double a) {
        return Math.signum(a);
    }

    public static int RANDBETWEEN(int a, int b) {
        return ThreadLocalRandom.current().nextInt(a, b + 1);
    }

    public static double MEDIAN(double... d) {
        Arrays.sort(d);
        int dl2 = d.length / 2;
        return d.length % 2 == 0 ? (d[dl2] + d[dl2 + 1]) / 2 : d[dl2];
    }
}
