package test.devices.osc;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.Date;

import static test.measure.Util.format;
import static test.measure.Util.get;
import static test.measure.Util.getCell;

public class CD extends OSC {
    protected double shiftLimit;
    protected double dLimit;
    protected double ch;

    public CD(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OSC/CDdataGS.properties";
        } else {
            dataFile = dataDir + "OSC/CDdata.properties";
        }
        ch = 20;
    }

    public CD(int slot, boolean isGS, double ch) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OSC/CDdataGS.properties";
        } else {
            dataFile = dataDir + "OSC/CDdata.properties";
        }
        this.ch = ch;
    }

    @Override
    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        super.test();
        OSA.setCenter(centerWL);
        OSA.setSpan(span);
        OSA.setVBW(VBW);
        OSA.WFILON();
        OSA.takeReferenceSpectrum(0);
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        while (counter != testCount) {
            switch (counter) {
                case 1:
                    OSCDTest(dLimit, ch, shiftLimit, "OSC-D IN - LINE2 OUT", "OSCDLN2", reader);
                    break;
                case 2:
                    upDTest(dLimit, "UPG2 IN - LINE2 OUT", "UP2LN2", reader);
                    break;
                case 3:
                    OSCDTest(dLimit, ch, shiftLimit, "LINE2 IN - OSC-D OUT", "LN2OSCD", reader);
                    break;
                case 4:
                    upDTest(dLimit, "LINE2 IN - UPG2 OUT", "LN2UP2", reader);
                    break;
                case 5:
                    upCTest(0.1, upLimit, "UPG1 IN - LINE1 OUT", "UP1LN1", reader);
                    break;
                case 6:
                    upCTest(mon1dLimit, mon1uLimit, "UPG1 IN - MON11 OUT", "UP1MON11", reader);
                    break;
                case 7:
                    upCTest(mon2dLimit, mon2uLimit, "LINE1 IN - MON12 OUT", "LN1MON12", reader);
                    break;
                case 8:
                    upCTest(0.1, upLimit, "LINE1 IN - UP1 OUT", "LN1UP1", reader);
                    break;
                case 9:
                    System.out.println("SFP - OSA");
                    reader.readLine();
                    OSA.setCenter(centerWL2);
                    OSA.setSpan(span2);
                    OSA.setTrace(1);
                    OSA.single();
                    SFPPeak = OSA.peakSearch() < 1.5 ? -1.5 : -OSA.peakSearch();
                    System.out.println(SFPPeak);
                    counter++;
                    break;
                case 10:
                    OSCCTest(0.1, OSCLimit,"OSC-C IN - LINE1 OUT", "OSCCLN1", reader);
                    break;
                case 11:
                    OSCCTest(mon1dLimit, mon1uLimit, "OSC-C IN - MON11 OUT", "OSCCMON11", reader);
                    break;
                case 12:
                    OSCCTest(0.1, OSCLimit,"LINE1 IN - OSC-C OUT", "LN1OSCC", reader);
                    break;
            }
        }
        OSA.APOFF();
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        if (!isGS) {
            getCell("H1", sheetFrom).setCellValue((int) ch);
        }
        getCell("E1", sheetFrom).setCellValue(get(protocolFile,"SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile,"PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile,"MSN"));
        getCell("E23", sheetFrom).setCellValue(get(protocolFile,"OSCDLN2W"));
        getCell("E24", sheetFrom).setCellValue(get(protocolFile,"LN2OSCDW"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile,"OSCCLN1"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile,"UP1LN1"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile,"LN1OSCC"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile,"LN1UP1"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile,"OSCCMON11"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile,"UP1MON11"));
        getCell("E35", sheetFrom).setCellValue(get(protocolFile,"LN1MON12"));
        getCell("E38", sheetFrom).setCellValue(get(protocolFile,"OSCDLN2"));
        getCell("E39", sheetFrom).setCellValue(get(protocolFile,"UP2LN2"));
        getCell("E40", sheetFrom).setCellValue(get(protocolFile,"LN2OSCD"));
        getCell("E41", sheetFrom).setCellValue(get(protocolFile,"LN2UP2"));
        getCell("B63", sheetFrom).setCellValue(format(new Date()));
        getCell("F63", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        shiftLimit = Double.parseDouble(get(dataFile,"SHIFT"));
        dLimit = Double.parseDouble(get(dataFile,"D"));
        if (!isGS) {
            pId = pId + (int) ch + "-01";
        }
    }
}
