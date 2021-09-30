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

public class OSC1C extends OSC {
    public OSC1C(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OSC/OSC1CdataGS.properties";
        } else {
            dataFile = dataDir + "OSC/OSC1Cdata.properties";
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
                    upCTest(0.1, upLimit, "UPG IN - LINE OUT", "UPLN", reader);
                    break;
                case 2:
                    upCTest(mon1dLimit, mon1uLimit, "UPG IN - MON1 OUT", "UPMON1", reader);
                    break;
                case 3:
                    upCTest(mon2dLimit, mon2uLimit, "LINE IN - MON2 OUT", "LNMON2", reader);
                    break;
                case 4:
                    upCTest(0.1, upLimit, "LINE IN - UPG OUT", "LNUP", reader);
                    break;
                case 5:
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
                case 6:
                    OSCCTest(0.1, OSCLimit,"OSC IN - LINE OUT", "OSCLN", reader);
                    break;
                case 7:
                    OSCCTest(mon1dLimit, mon1uLimit, "OSC IN - MON1 OUT", "OSCMON1", reader);
                    break;
                case 8:
                    OSCCTest(0.1, OSCLimit,"LINE IN - OSC OUT", "LNOSC", reader);
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
        getCell("E23", sheetFrom).setCellValue(get(protocolFile,"OSCLN"));
        getCell("E24", sheetFrom).setCellValue(get(protocolFile,"UPLN"));
        getCell("E25", sheetFrom).setCellValue(get(protocolFile,"LNOSC"));
        getCell("E26", sheetFrom).setCellValue(get(protocolFile,"LNUP"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile,"OSCMON1"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile,"UPMON1"));
        getCell("E31", sheetFrom).setCellValue(get(protocolFile,"LNMON2"));
        getCell("B47", sheetFrom).setCellValue(format(new Date()));
        getCell("F47", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
