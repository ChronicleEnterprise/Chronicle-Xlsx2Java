package software.chronicle.xlsx2java;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.IOException;
import java.util.Arrays;

import static org.junit.Assert.assertEquals;
import static software.chronicle.xlsx2java.XLSXReader.depends;

/**
 * Created by peter on 19/07/2017.
 */
public class XLSXReaderTest {

    @Test
    public void testDepends() {
        assertEquals("[K19, W14, X14, Z4]", Arrays.toString(depends("IF($K$19>0,CEILING(W14,0.00001),X14+Z4)")));
        assertEquals("[K3, K15]", Arrays.toString(depends("MEDIAN(K3:K15)")));
    }

    @Test
    public void testProcess() throws IOException, InvalidFormatException {
        String name = "example1.xlsx";
        XLSXReader reader = new XLSXReader(name);
        String process = reader.process();
        System.out.println(process);
        assertEquals(
                "    double B3;\n" +
                        "    double C3;\n" +
                        "    double B4;\n" +
                        "    double B5;\n" +
                        "    double B6;\n" +
                        "    double B7;\n" +
                        "    double B8;\n" +
                        "    double B9;\n" +
                        "    public void reset() {\n" +
                        "        B3 = 1.0;\n" +
                        "        C3 = 1.5;\n" +
                        "        B4 = Double.NaN;\n" +
                        "        B5 = Double.NaN;\n" +
                        "        B6 = Double.NaN;\n" +
                        "        B7 = Double.NaN;\n" +
                        "        B8 = Double.NaN;\n" +
                        "        B9 = Double.NaN;\n" +
                        "    }\n" +
                        "    public void onePass() {\n" +
                        "        B4 = MIN(B3,C3);\n" +
                        "        B5 = MAX(B3,C3);\n" +
                        "        B6 = IF(B3>C3, B3, C3);\n" +
                        "        B7 = ROUND(C3, 0);\n" +
                        "        B8 = CEILING(C3, 0.1);\n" +
                        "        B9 = FLOOR(C3, 0.1);\n" +
                        "    }\n" +
                        "    public void twoPass() {\n" +
                        "    }\n", process);

    }
}