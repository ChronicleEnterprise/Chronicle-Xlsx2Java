package software.chronicle.xlsx2java;

import net.openhft.chronicle.core.io.IOTools;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.StringWriter;
import java.net.URL;
import java.util.*;

/**
 * Created by peter on 18/07/2017.
 */
public class XLSXReader {
    private final String name;
    private String indent = "    ";
    private Map<String, String> values = new LinkedHashMap<>();
    private Set<String> used = new LinkedHashSet<>();
    private Set<String> overridden = new LinkedHashSet<>();
    private Set<String> isFormula = new LinkedHashSet<>();
    private Map<String, CellValue> cells = new LinkedHashMap<>();

    public XLSXReader(String name) {
        this.name = name;
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        String name = args[0];
        XLSXReader reader = new XLSXReader(name);
//        IntStream.rangeClosed(3, 17).forEach(i -> reader.override("K" + i));
//        reader.override("K21,K22,K23");
//        reader.override("S4,S5,S6,S7,S8");
        System.out.println(reader.process());
    }

    public static String[] depends(String formula) {
        List<String> ids = new ArrayList<>();
        StringBuilder sb = new StringBuilder();
        boolean letter = false;
        boolean digit = false;
        for (int i = 0; i <= formula.length(); i++) {
            char ch = i < formula.length() ? formula.charAt(i) : ' ';
            if (ch == '$')
                continue;
            if ('A' <= ch && ch <= 'Z') {
                letter = true;
                sb.append(ch);
                continue;
            }
            if ('0' <= ch && ch <= '9') {
                digit = true;
                sb.append(ch);
                continue;
            }
            if (letter && digit) {
                ids.add(sb.toString());
            }
            sb.setLength(0);
            letter = digit = false;
        }
        return ids.toArray(new String[ids.size()]);
    }

    private void override(String cells) {
        Collections.addAll(overridden, cells.toUpperCase().split(","));
    }

    public String process() throws IOException, InvalidFormatException {
        URL url = IOTools.urlFor(name);
        Workbook wb = new XSSFWorkbook(new File(url.getFile()));

        for (Sheet sheet : wb) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String cr = cell.getAddress().formatAsString();
                    switch (cell.getCellTypeEnum()) {
                        case _NONE:
                            break;

                        case NUMERIC:
                            values.put(cr, "double " + cr + " = " + cell.getNumericCellValue() + ";");
                            break;

                        case STRING:
                            values.put(cr, "String " + cr + " = \"" + cell.getStringCellValue() + "\";");
                            break;

                        case FORMULA:
                            if (overridden.contains(cr)) {
                                values.put(cr, "double " + cr + " = Double.NaN;");

                            } else {
                                CellValue cellValue = new CellValue(cr, cell.getCellFormula().replace("$", ""));
                                cells.put(cr, cellValue);
                                isFormula.add(cr);
                            }
                            break;

                        case BLANK:
                            break;

                        case BOOLEAN:
                            break;

                        case ERROR:
                            break;
                    }
                }
            }
        }


        for (int i = 0; i < 50; i++) {
            for (CellValue cell : cells.values()) {
                int r = cell.rank;
                boolean changed = false;
                for (String s : cell.depends) {
                    CellValue cell2 = cells.get(s);
                    if (cell2 == null) {
                        continue;
                    }
                    int r2 = cell2.rank;
                    if (r2 >= r) {
                        r = r2 + 1;
                        changed = true;
                    }
                }
                if (changed)
                    cell.rank = r;
            }
        }

        used.addAll(overridden);
        List<CellValue> cellValueList = new ArrayList<>(cells.values());
        cellValueList.sort(Comparator.comparing(c -> c.rank));
        for (CellValue cellValue : cellValueList) {
            Collections.addAll(used, cellValue.depends);
        }

        StringWriter out = new StringWriter();

        // get the field values.
        for (String s : used) {
            if (isFormula.contains(s))
                continue;
            String x = values.get(s);
            if (x == null) {
                if (!overridden.contains(s))
                    System.err.println("Missing '" + s + "'");
            } else {
                out.append(indent).append(x.split(" = ", 2)[0]).append(";\n");
            }
        }
        for (CellValue cell : cellValueList) {
            out.append(indent).append("double " + cell.cellId + ";\n");
        }

        out.append(indent).append("public void reset() {\n");
        // reset the field values
        for (String s : used) {
            if (isFormula.contains(s) || overridden.contains(s))
                continue;
            String x = values.get(s);
            if (x == null)
                System.err.println("Missing '" + s + "'");
            else {
                out.append(indent).append(indent).append(x.split(" ", 2)[1]).append("\n");
            }
        }
        for (CellValue cell : cellValueList) {
            int r = cell.rank;
            if (r < 50)
                out.append(indent).append(indent).append(cell.cellId + " = Double.NaN;\n");
        }
        out.append(indent).append("}\n");

        out.append(indent).append("public void onePass() {\n");
        // non circular
        for (CellValue cell : cellValueList) {
            int r = cell.rank;
            if (r < 50)
                out.append(indent).append(indent).append(cell.cellId + " = " + cell.formula + ";\n");
        }
        out.append(indent).append("}\n");

        // circular
        out.append(indent).append("public void twoPass() {\n");

        for (CellValue cell : cellValueList) {
            int r = cell.rank;
            if (r >= 50)
                out.append(indent).append(indent).append(cell.cellId + " = " + cell.formula + ";\n");
        }
        out.append(indent).append("}\n");

        return out.toString();
    }

    static class CellValue {
        String cellId;
        String formula;
        String[] depends;
        int rank;

        public CellValue(String cellId, String formula) {
            this.cellId = cellId;
            this.formula = formula
                    .replace("K3:K15", "K3,K4,K5,K6,K7,K8,K9,K10,K11,K12,K13,K14,K15")
                    .replaceAll("(?![<>])=", "==");
            depends = depends(this.formula);
        }
    }
}
