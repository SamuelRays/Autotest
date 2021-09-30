package test.devices.roadm;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class ROADM21 extends ROADM {
    private double indropLimit;
    private double inroutLimit;
    private double rinoutLimit;
    private double addoutLimit;
    private double addmonLimit;
    private double rinmonLimit;
    private double dev9Limit;
    private double dev15Limit;

    public ROADM21(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "ROADM/ROADM21dataGS.properties";
            isKama = true;
        } else {
            dataFile = dataDir + "ROADM/ROADM21data.properties";
        }
    }

    @Override
    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        super.test();
        OSA.setCenter(centerWL);
        OSA.setSpan(span2);
        OSA.setVBW(VBW3);
        OSA.takeReferenceSpectrum(3);
        OSA.setSpan(span);
        OSA.setVBW(VBW);
        OSA.takeReferenceSpectrum(0);
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        while (counter != testCount) {
            switch (counter) {
                case 1:
                    System.out.println("IN - DROP");
                    reader.readLine();
                    if (checkThresholdsWriteResults(0.1, indropLimit, "INDROP", protocolFile, OSA)) {
                        counter++;
                    }
                    break;
                case 2:
                    System.out.println("IN - ROUT");
                    reader.readLine();
                    if (checkThresholdsWriteResults(0.1, inroutLimit, "INROUT", protocolFile, OSA)) {
                        counter++;
                    }
                    break;
                case 3:
                    lossFlatTest(1, "RIN - OUT", 0.1, rinoutLimit, "RINOUT", reader);
                    break;
                case 4:
                    lossFlatTest(2, "ADD - OUT", 0.1, addoutLimit, "ADDOUT", reader);
                    set(protocolFile,"FLATNESS", String.valueOf(round(currentFlatness, 10)));
                    break;
                case 5:
                    lossTest(2, "ADD - MON OUT", 20, addmonLimit, "ADDMON", reader);
                    break;
                case 6:
                    lossTest(1, "RIN - MON OUT", 20, rinmonLimit, "RINMON", reader);
                    break;
                case 7:
                    System.out.println("RIN - OUT");
                    reader.readLine();
                    deviationTest(true, 1, 9, dev9Limit);
                    break;
                case 8:
                    deviationTest(false, 1, 15, dev15Limit);
                    break;
                case 9:
                    allChannelTest(100);
            }
        }
        volga.logout();
        OSA.logout();
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        int st = isGS ? 23 : 22;
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        getCell("E1", sheetFrom).setCellValue(get(protocolFile,"SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile,"PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile,"MSN"));
        getCell("E" + st, sheetFrom).setCellValue(get(protocolFile,"INDROP"));
        getCell("E" + (st + 1), sheetFrom).setCellValue(get(protocolFile,"INROUT"));
        getCell("E" + (st + 2), sheetFrom).setCellValue(get(protocolFile,"RINOUT"));
        getCell("E" + (st + 3), sheetFrom).setCellValue(get(protocolFile,"ADDOUT"));
        getCell("E" + (st + 4), sheetFrom).setCellValue(get(protocolFile,"ADDMON"));
        getCell("E" + (st + 5), sheetFrom).setCellValue(get(protocolFile,"RINMON"));
        getCell("E" + (st + 10), sheetFrom).setCellValue(get(protocolFile,"WLMAX"));
        getCell("E" + (st + 11), sheetFrom).setCellValue(get(protocolFile,"WLMIN"));
        getCell("E" + (st + 13), sheetFrom).setCellValue(get(protocolFile,"SHIFT"));
        getCell("E" + (st + 16), sheetFrom).setCellValue(get(protocolFile,"FLATNESS"));
        getCell("E" + (st + 17), sheetFrom).setCellValue(get(protocolFile,"DEVIATION9"));
        getCell("E" + (st + 18), sheetFrom).setCellValue(get(protocolFile,"DEVIATION15"));
        getCell("B" + (st + 37), sheetFrom).setCellValue(format(new Date()));
        getCell("E" + (st + 37), sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        indropLimit = Double.parseDouble(get(dataFile,"INDROP"));
        inroutLimit = Double.parseDouble(get(dataFile,"INROUT"));
        rinoutLimit = Double.parseDouble(get(dataFile,"RINOUT"));
        addoutLimit = Double.parseDouble(get(dataFile,"ADDOUT"));
        addmonLimit = Double.parseDouble(get(dataFile,"ADDMON"));
        rinmonLimit = Double.parseDouble(get(dataFile,"RINMON"));
        dev9Limit = Double.parseDouble(get(dataFile,"DEVIATION9"));
        dev15Limit = Double.parseDouble(get(dataFile,"DEVIATION15"));
    }
}
