package software.chronicle.xlsx2java;

import org.junit.Test;

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
        assertEquals("", Arrays.toString(depends("MEDIAN(K3:K15)")));
    }
}