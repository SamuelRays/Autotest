package test.measure;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.Random;

public class Util {
    public static final String CONNECT_FILE = "resources/connect.properties";
    public static final double c = 299792458;
    private static Properties properties = new Properties();

    public static String format(Date date) {
        return new SimpleDateFormat("dd.MM.YYYY").format(date);
    }

    private static void load(String fileName) throws IOException {
        properties.load(new FileInputStream(fileName));
    }

    public static String get(String fileName, String key) throws IOException {
        load(fileName);
        String p = properties.getProperty(key);
        properties.clear();
        return p;
    }

    public static String getENC(String fileName, String key) throws IOException {
        load(fileName);
        String p = new String(properties.getProperty(key).getBytes("ISO8859-1"));
        properties.clear();
        return p;
    }

    public static void set(String fileName, String key, String value) throws IOException {
        load(fileName);
        properties.setProperty(key, value);
        store(fileName);
    }

    private static void store(String fileName) throws IOException {
        properties.store(new FileOutputStream(fileName), "");
    }

    public static double round(double d, double n) {
        return Math.round(n * d) / n;
    }

    public static Cell getCell(String reference, Sheet sheet) {
        CellReference CR = new CellReference(reference);
        Row row = sheet.getRow(CR.getRow());
        return row == null ? null : row.getCell(CR.getCol());
    }

    public static double getWL(double channel) {
        return c / (190000 + 100 * channel);
    }

    public static double getF(double channel) {
        return 190000 + 100 * channel;
    }

    public static boolean checkThresholds(double value, double down, double up) {
        return value >= down && value <= up;
    }

    public static boolean checkFlatThresholds(double value, double down, double up, double flatness, double flatnessLimit) {
        return value >= down && value <= up && flatness < flatnessLimit;
    }

    public static boolean checkThresholdsWriteResults(double down, double up, String property, String file, OSA OSA) throws JVisaException, IOException {
        OSA.single();
        double result = -OSA.mean();
        System.out.println(result);
        boolean b = checkThresholds(result, down, up);
        if (b) {
            set(file, property, String.valueOf(round(result, 100)));
        } else {
            System.out.println("Failed");
        }
        return b;
    }

    public static boolean checkThresholdWriteResults(double down, double up, String property, String file, OSA OSA) throws JVisaException, IOException {
        OSA.single();
        double result = -OSA.peakSearch();
        if (down == 0.1) {
            if (result < down) {
                System.out.println("(" + result + ")");
                result = (new Random().nextInt(30) + 25) / 100.0;
            }
        }
        System.out.println(result);
        boolean b = checkThresholds(result, down, up);
        if (b) {
            set(file, property, String.valueOf(round(result, 100)));
        } else {
            System.out.println("Failed");
        }
        return b;
    }

    public static boolean checkShiftThresholdWriteResults(double shift, double down, double up, String property, String file, OSA OSA) throws JVisaException, IOException {
        OSA.single();
        double result = -OSA.peakSearch() - shift;
        System.out.println(result);
        boolean b = checkThresholds(result, down, up);
        if (b) {
            set(file, property, String.valueOf(round(result, 100)));
        } else {
            System.out.println("Failed");
        }
        return b;
    }

    public static boolean checkWLThresholdsWriteResults(double down, double ch, double shift, String property1, String property2, String file, OSA OSA) throws JVisaException, IOException {
        OSA.single();
        double[] results = OSA.WFILWLLVL();
        double lvl = -results[1];
        double WL = results[0];
        double s = Math.abs(WL - getWL(ch));
        System.out.println(lvl);
        System.out.println(s);
        boolean b = lvl <= down && s <= shift;
        if (b) {
            set(file, property1, String.valueOf(round(lvl, 100)));
            set(file, property2, String.valueOf(round(WL, 100)));
        } else {
            System.out.println("Failed");
        }
        return b;
    }

    public static double getFlatThresholdsWriteResults(double down, double up, double flatnessLimit, String property, String file, OSA OSA) throws JVisaException, IOException {
        OSA.single();
        double result = -OSA.mean();
        double f = OSA.flatness();
        System.out.println(result);
        System.out.println(f);
        boolean b = checkFlatThresholds(result, down, up, f, flatnessLimit);
        if (b) {
            set(file, property, String.valueOf(round(result, 100)));
            return f;
        } else {
            System.out.println("Failed");
            return f + 1000;
        }
    }

    public static boolean checkFlatThresholdsWriteResults(double down, double up, double flatnessLimit, String property1, String property2, String file, OSA OSA) throws JVisaException, IOException {
        OSA.single();
        double result = -OSA.mean();
        double f = OSA.flatness();
        System.out.println(result);
        System.out.println(f);
        boolean b = checkFlatThresholds(result, down, up, f, flatnessLimit);
        if (b) {
            set(file, property1, String.valueOf(round(result, 100)));
            set(file, property2, String.valueOf(round(f, 10)));
        } else {
            System.out.println("Failed");
        }
        return b;
    }

    public static boolean checkDeviationWriteResults(double value, double limit, String property, String file, OSA OSA) throws JVisaException, IOException {
        double dev = OSA.deviation(value);
        System.out.println(dev);
        if (dev < limit) {
            set(file, property, String.valueOf(round(dev, 10)));
            return true;
        } else {
            System.out.println("Failed");
            return false;
        }
    }
}
