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

public class ROADM91 extends ROADM {
    private double indropLimit;
    private double inmonLimit;
    private double[] addLimits = new double[9];
    private double mondLimit;
    private double monuLimit;
    private double dev9Limit;
    private double dev20Limit;

    public ROADM91(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "ROADM/ROADM91dataGS.properties";
            isKama = true;
        } else {
            dataFile = dataDir + "ROADM/ROADM91data.properties";
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
                    System.out.println("IN - MON IN");
                    reader.readLine();
                    if (checkThresholdsWriteResults(0.1, inmonLimit, "INRMON", protocolFile, OSA)) {
                        counter++;
                    }
                    break;
                case 3:
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                case 9:
                case 10:
                case 11:
                    lossFlatTest(counter - 2, "ADD" + (counter - 2) + " - OUT", 0.1, addLimits[counter - 3], "ADD" + (counter - 2) + "OUT", reader);
                    break;
                case 12:
                    deviationTest(true, 9, 9, dev9Limit);
                    break;
                case 13:
                    deviationTest(false, 9, 20, dev20Limit);
                    break;
                case 14:
                    lossTest(9, "ADD9 - MON OUT", mondLimit, monuLimit, "MON", reader);
                    break;
                case 15:
                    System.out.println("ADD1 - OUT");
                    reader.readLine();
                    allChannelTest(1000);
            }
        }
        set(protocolFile,"FLATNESS", String.valueOf(round(currentFlatness, 10)));
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
        getCell("E" + (st + 1), sheetFrom).setCellValue(get(protocolFile,"INMON"));
        getCell("E" + (st + 2), sheetFrom).setCellValue(get(protocolFile,"ADD1OUT"));
        getCell("E" + (st + 3), sheetFrom).setCellValue(get(protocolFile,"ADD2OUT"));
        getCell("E" + (st + 4), sheetFrom).setCellValue(get(protocolFile,"ADD3OUT"));
        getCell("E" + (st + 5), sheetFrom).setCellValue(get(protocolFile,"ADD4OUT"));
        getCell("E" + (st + 6), sheetFrom).setCellValue(get(protocolFile,"ADD5OUT"));
        getCell("E" + (st + 7), sheetFrom).setCellValue(get(protocolFile,"ADD6OUT"));
        getCell("E" + (st + 8), sheetFrom).setCellValue(get(protocolFile,"ADD7OUT"));
        getCell("E" + (st + 9), sheetFrom).setCellValue(get(protocolFile,"ADD8OUT"));
        getCell("E" + (st + 10), sheetFrom).setCellValue(get(protocolFile,"ADD9OUT"));
        getCell("E" + (st + 13), sheetFrom).setCellValue(get(protocolFile,"MON"));
        getCell("E" + (st + 18), sheetFrom).setCellValue(get(protocolFile,"WLMAX"));
        getCell("E" + (st + 19), sheetFrom).setCellValue(get(protocolFile,"WLMIN"));
        getCell("E" + (st + 21), sheetFrom).setCellValue(get(protocolFile,"SHIFT"));
        getCell("E" + (st + 24), sheetFrom).setCellValue(get(protocolFile,"FLATNESS"));
        getCell("E" + (st + 25), sheetFrom).setCellValue(get(protocolFile,"DEVIATION9"));
        getCell("E" + (st + 26), sheetFrom).setCellValue(get(protocolFile,"DEVIATION20"));
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
        inmonLimit = Double.parseDouble(get(dataFile,"INMON"));
        addLimits[0] = Double.parseDouble(get(dataFile,"ADD1OUT"));
        addLimits[1] = Double.parseDouble(get(dataFile,"ADD2OUT"));
        addLimits[2] = Double.parseDouble(get(dataFile,"ADD3OUT"));
        addLimits[3] = Double.parseDouble(get(dataFile,"ADD4OUT"));
        addLimits[4] = Double.parseDouble(get(dataFile,"ADD5OUT"));
        addLimits[5] = Double.parseDouble(get(dataFile,"ADD6OUT"));
        addLimits[6] = Double.parseDouble(get(dataFile,"ADD7OUT"));
        addLimits[7] = Double.parseDouble(get(dataFile,"ADD8OUT"));
        addLimits[8] = Double.parseDouble(get(dataFile,"ADD9OUT"));
        mondLimit = Double.parseDouble(get(dataFile,"MOND"));
        monuLimit = Double.parseDouble(get(dataFile,"MONU"));
        dev9Limit = Double.parseDouble(get(dataFile,"DEVIATION9"));
        dev20Limit = Double.parseDouble(get(dataFile,"DEVIATION20"));
    }
}
