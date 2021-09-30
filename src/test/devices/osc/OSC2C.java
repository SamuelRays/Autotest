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

public class OSC2C extends OSC {
    public OSC2C(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OSC/OSC2CdataGS.properties";
        } else {
            dataFile = dataDir + "OSC/OSC2Cdata.properties";
        }
    }

    @Override
    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        super.test();
        OSA.setCenter(centerWL);
        OSA.setSpan(span);
        OSA.setVBW(VBW);
        OSA.takeReferenceSpectrum(0);
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        while (counter != testCount) {
            switch (counter) {
                case 1:
                    upCTest(0.1, upLimit, "UPG1 IN - LINE1 OUT", "UP1LN1", reader);
                    break;
                case 2:
                    upCTest(mon1dLimit, mon1uLimit, "UPG1 IN - MON11 OUT", "UP1MON11", reader);
                    break;
                case 3:
                    upCTest(mon2dLimit, mon2uLimit, "LINE1 IN - MON12 OUT", "LN1MON12", reader);
                    break;
                case 4:
                    upCTest(0.1, upLimit, "LINE1 IN - UPG1 OUT", "LN1UP1", reader);
                    break;
                case 5:
                    upCTest(0.1, upLimit, "UPG2 IN - LINE2 OUT", "UP2LN2", reader);
                    break;
                case 6:
                    upCTest(mon1dLimit, mon1uLimit, "UPG2 IN - MON21 OUT", "UP2MON21", reader);
                    break;
                case 7:
                    upCTest(mon2dLimit, mon2uLimit, "LINE2 IN - MON22 OUT", "LN2MON22", reader);
                    break;
                case 8:
                    upCTest(0.1, upLimit, "LINE2 IN - UPG2 OUT", "LN2UP2", reader);
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
                    OSCCTest(0.1, OSCLimit,"OSC1 IN - LINE1 OUT", "OSC1LN1", reader);
                    break;
                case 11:
                    OSCCTest(mon1dLimit, mon1uLimit, "OSC1 IN - MON11 OUT", "OSC1MON11", reader);
                    break;
                case 12:
                    OSCCTest(0.1, OSCLimit,"LINE1 IN - OSC1 OUT", "LN1OSC1", reader);
                    break;
                case 13:
                    OSCCTest(0.1, OSCLimit,"OSC2 IN - LINE2 OUT", "OSC2LN2", reader);
                    break;
                case 14:
                    OSCCTest(mon1dLimit, mon1uLimit, "OSC2 IN - MON21 OUT", "OSC2MON21", reader);
                    break;
                case 15:
                    OSCCTest(0.1, OSCLimit,"LINE2 IN - OSC2 OUT", "LN2OSC2", reader);
                    break;
            }
        }
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        getCell("E1", sheetFrom).setCellValue(get(protocolFile,"SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile,"PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile,"MSN"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile,"OSC1LN1"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile,"UP1LN1"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile,"LN1OSC1"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile,"LN1UP1"));
        getCell("E31", sheetFrom).setCellValue(get(protocolFile,"OSC2LN2"));
        getCell("E32", sheetFrom).setCellValue(get(protocolFile,"UP2LN2"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile,"LN2OSC2"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile,"LN2UP2"));
        getCell("E37", sheetFrom).setCellValue(get(protocolFile,"OSC1MON11"));
        getCell("E38", sheetFrom).setCellValue(get(protocolFile,"UP1MON11"));
        getCell("E39", sheetFrom).setCellValue(get(protocolFile,"LN1MON12"));
        getCell("E40", sheetFrom).setCellValue(get(protocolFile,"OSC2MON21"));
        getCell("E41", sheetFrom).setCellValue(get(protocolFile,"UP2MON21"));
        getCell("E42", sheetFrom).setCellValue(get(protocolFile,"LN2MON22"));
        getCell("B66", sheetFrom).setCellValue(format(new Date()));
        getCell("E66", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
