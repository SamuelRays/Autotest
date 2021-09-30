package test.devices;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;
import test.measure.OSA;
import test.measure.Util;
import volga.Crate;
import volga.Volga;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import static test.measure.Util.*;

public abstract class Device {
    protected static String dataDir;
    protected static String protocolDir;
    protected int counter = 1;
    protected String pId;
    protected String SN;
    protected String PN;
    protected boolean isMSN;
    protected boolean isGS;
    protected int testCount;
    protected OSA OSA;
    protected double centerWL;
    protected double span;
    protected String VBW;
    protected Volga volga;
    protected int slot;
    protected String protocolFile;
    protected String mapFile;
    protected String sampleFile;
    protected String dir;
    protected String dataFile;
    protected String engineer;

    public Device(int slot, boolean isGS) throws IOException {
        this.slot = slot;
        this.isGS = isGS;
        OSA = new OSA(get(CONNECT_FILE,"OSAURL"));
        volga = new Volga(Crate.valueOf(get(CONNECT_FILE,"crateType")), get(CONNECT_FILE,"VolgaURL"));
        dataDir = get(CONNECT_FILE, "DATADIR");
        protocolDir = get(CONNECT_FILE, "PROTOCOLDIR");
        engineer = getENC(CONNECT_FILE, "ENGINEER");
    }

    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        prepare();
    }

    protected void setSNs() throws IOException, TransformerException, ParserConfigurationException, SAXException {
        PN = volga.CUget("Slot" + slot + "Name");
        if (isGS) {
            PN = PN.substring(0, PN.length() - 3);
        }
        InputStream inputStreamFrom = new FileInputStream(mapFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        String MSN = null;
        int i = 1;
        while (true) {
            Cell cell = getCell("B" + i, sheetFrom);
            if (cell == null) {
                i++;
                continue;
            }
            String c = cell.getStringCellValue();
            if (c.equals(PN)) {
                SN = getCell("A" +i, sheetFrom).getStringCellValue();
                if (isMSN) {
                    Cell MSNCell = getCell("C" +i, sheetFrom);
                    MSNCell.setCellType(CellType.STRING);
                    MSN = MSNCell.getStringCellValue();
                }
                break;
            }
            i++;
        }
        workbookFrom.close();
        System.out.println(SN);
        set(protocolFile,"PN", PN);
        set(protocolFile,"SN", SN);
        if (isMSN) {
            set(protocolFile,"MSN", MSN);
        }
    }

    protected abstract void setPower() throws TransformerException, ParserConfigurationException, IOException, SAXException, InterruptedException;

    protected void prepare() throws InterruptedException, ParserConfigurationException, TransformerException, SAXException, IOException, JVisaException {
        readData();
        volga.connect();
        setPower();
        setSNs();
        OSA.connect();
    }

    public void createProtocol() throws IOException {
        if (protocolFile == null) {
            readData();
        }
    }

    protected void readData() throws IOException {
        protocolFile = dataDir + get(dataFile,"PFILE");
        dir = protocolDir + get(dataFile,"DIR");
        sampleFile = dir + "_sample.xlsx";
        mapFile = dir + "_mapping.xlsx";
        pId = get(dataFile,"PID");
        testCount = Integer.parseInt(get(dataFile,"TESTCOUNT"));
        span = Double.parseDouble(get(dataFile,"SPAN"));
        centerWL = Double.parseDouble(get(dataFile,"CENTER"));
        VBW = get(dataFile,"VBW");
    }

    public void setParam(String param, String value) throws TransformerException, ParserConfigurationException {
        volga.setSlotParam(slot, param, value);
    }

    public String getParam(String param) throws TransformerException, ParserConfigurationException, IOException, SAXException {
        return volga.getSlotParam(slot, param);
    }

    public String getpId() {
        return pId;
    }

    public String getSN() {
        return SN;
    }

    public void setSN(String SN) {
        this.SN = SN;
    }

    public String getPN() {
        return PN;
    }

    public void setPN(String PN) {
        this.PN = PN;
    }

    public boolean isMSN() {
        return isMSN;
    }

    public boolean isGS() {
        return isGS;
    }

    public int getTestCount() {
        return testCount;
    }

    public OSA getOSA() {
        return OSA;
    }

    public Volga getVolga() {
        return volga;
    }
}
