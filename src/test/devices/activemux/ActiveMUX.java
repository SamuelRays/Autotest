package test.devices.activemux;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;
import test.devices.ActiveDevice;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import static test.measure.Util.*;

public abstract class ActiveMUX extends ActiveDevice {
    protected double att1Limit;
    protected double att2Limit;
    protected double att3Limit;
    protected double shiftLimit;
    protected double dev9Limit;
    protected double dev15Limit;
    protected double monLimit;
    protected double currentDev9;
    protected double currentDev15;
    protected boolean isPlus;

    public ActiveMUX(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isMSN = true;
    }

    public ActiveMUX(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS, isKama);
        isMSN = true;
    }

    @Override
    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        super.test();
        OSA.setCenter(centerWL);
        OSA.setSpan(span);
        OSA.setVBW(VBW);
        OSA.takeReferenceSpectrum(0);
        OSA.WFILON();
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        MUXON();
        setAllChAtt(0);
        do {
            System.out.println("MON");
            reader.readLine();
        } while (!checkThresholdWriteResults(10, monLimit, "MON", protocolFile, OSA));
        for (int i = 0; i < 3; i++) {
            double att = i == 0 ? 0 : i == 1 ? 9 : 15;
            if (i == 1) {
                setAllChAtt(9);
            } else if (i == 2) {
                setAllChAtt(15);
            }
            while (counter != testCount) {
                channelTest(counter + 20, att, reader);
            }
            counter = 1;
        }
        set(protocolFile, "DEV9", String.valueOf(round(currentDev9, 100)));
        set(protocolFile, "DEV15", String.valueOf(round(currentDev15, 100)));
        OSA.APOFF();
        setAllChAtt(0);
        MUXOFF();
    }

    protected void channelTest(int ch, double att, BufferedReader reader) throws IOException, JVisaException {
        String message = isPlus ? "CH" + ch + "+" : "CH" + ch;
        String property = att == 0 ? "" : att == 9 ? "9" : "15";
        double channel = isPlus ? ch + 0.5 : ch;
        double limit = att == 0 ? att1Limit : att == 9 ? att2Limit : att3Limit;
        System.out.println(message);
        reader.readLine();
        OSA.single();
        double[] results = OSA.WFILWLLVL();
        double lvl = -results[1];
        double wl = results[0];
        double s = Math.abs(wl - getWL(channel));
        System.out.println(lvl);
        System.out.println(s);
        if (lvl <= limit && s <= shiftLimit) {
            if (att != 0) {
                double prev = Double.parseDouble(get(protocolFile, "CH" + ch));
                double dev = Math.abs(lvl - prev - att);
                System.out.println(dev);
                if (att == 9) {
                    if (dev > dev9Limit) {
                        System.out.println("Failed");
                        return;
                    }
                    currentDev9 = dev > currentDev9 ? dev : currentDev9;
                } else if (att == 15) {
                    if (dev > dev15Limit) {
                        System.out.println("Failed");
                        return;
                    }
                    currentDev15 = dev > currentDev15 ? dev : currentDev15;
                }
            }
            set(protocolFile, "CH" + ch + property, String.valueOf(round(lvl, 100)));
            if (att == 0) {
                set(protocolFile, "CH" + ch + "W", String.valueOf(round(wl, 100)));
            }
            counter++;
        } else {
            System.out.println("Failed");
        }
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        int st = isGS ? 24 : 22;
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        getCell("E1", sheetFrom).setCellValue(get(protocolFile,"SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile,"PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile,"MSN"));
        for (int i = st; i < st + 40; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"CH" + (i - st + 21) + "W"));
        }
        for (int i = st + 42; i < st + 82; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"CH" + (i - st - 21)));
        }
        for (int i = st + 85; i < st + 125; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile, "CH" + (i - st - 64) + "9"));
        }
        for (int i = st + 129; i < st + 169; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"CH" + (i - st - 108) + "15"));
        }
        getCell("E" + (st + 82), sheetFrom).setCellValue(get(protocolFile,"MON"));
        getCell("E" + (st + 126), sheetFrom).setCellValue(get(protocolFile,"DEV9"));
        getCell("E" + (st + 170), sheetFrom).setCellValue(get(protocolFile,"DEV15"));
        getCell("B" + (st + 223), sheetFrom).setCellValue(format(new Date()));
        getCell("E" + (st + 223), sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }

    protected void MUXON() throws TransformerException, ParserConfigurationException, InterruptedException {
        setAllChState(1);
        setAllChState(2);
    }

    protected void MUXOFF() throws TransformerException, ParserConfigurationException, InterruptedException {
        setAllChState(1);
        setAllChState(0);
    }

    protected void setAllChState(int state) throws TransformerException, ParserConfigurationException, InterruptedException {
        if (isKama) {
            setParam("ChAll.AdmState.Set", "" + state);
        } else {
            setParam("ChAll.Port.State", "" + state);
            TimeUnit.SECONDS.sleep(3);
        }
    }

    protected void setAllChAtt(double att) throws TransformerException, ParserConfigurationException, InterruptedException {
        if (isKama) {
            setParam("ChAll.VOA.Att.Set", "" + att);
        } else {
            setParam("VOAattSetAll", "" + att);
            TimeUnit.SECONDS.sleep(3);
        }
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        att1Limit = Double.parseDouble(get(dataFile, "ATT1"));
        att2Limit = Double.parseDouble(get(dataFile, "ATT2"));
        att3Limit = Double.parseDouble(get(dataFile, "ATT3"));
        shiftLimit = Double.parseDouble(get(dataFile, "SHIFT"));
        dev9Limit = Double.parseDouble(get(dataFile, "DEV9"));
        dev15Limit = Double.parseDouble(get(dataFile, "DEV15"));
        monLimit = Double.parseDouble(get(dataFile, "MON"));
    }
}
