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

public class ROADM41 extends ROADM {
    private double[] addLimits = new double[4];
    private double mondLimit;
    private double monuLimit;
    private double dev9Limit;
    private double dev15Limit;

    public ROADM41(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "ROADM/ROADM41dataGS.properties";
            isKama = true;
        } else {
            dataFile = dataDir + "ROADM/ROADM41data.properties";
            isKama = true;
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
                case 2:
                case 3:
                case 4:
                    lossFlatTest(counter, "ADD" + counter + " - OUT", 0.1, addLimits[counter - 1], "ADD" + counter + "OUT", reader);
                    break;
                case 5:
                    deviationTest(true, 4, 9, dev9Limit);
                    break;
                case 6:
                    deviationTest(false, 4, 15, dev15Limit);
                    break;
                case 7:
                    lossTest(4, "ADD4 - MON", mondLimit, monuLimit, "MON", reader);
                    break;
                case 8:
                    System.out.println("ADD1 - OUT");
                    reader.readLine();
                    allChannelTest(100);
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
        getCell("E" + st, sheetFrom).setCellValue(get(protocolFile,"ADD1OUT"));
        getCell("E" + (st + 1), sheetFrom).setCellValue(get(protocolFile,"ADD2OUT"));
        getCell("E" + (st + 2), sheetFrom).setCellValue(get(protocolFile,"ADD3OUT"));
        getCell("E" + (st + 3), sheetFrom).setCellValue(get(protocolFile,"ADD4OUT"));
        getCell("E" + (st + 6), sheetFrom).setCellValue(get(protocolFile,"MON"));
        getCell("E" + (st + 11), sheetFrom).setCellValue(get(protocolFile,"WLMAX"));
        getCell("E" + (st + 12), sheetFrom).setCellValue(get(protocolFile,"WLMIN"));
        getCell("E" + (st + 14), sheetFrom).setCellValue(get(protocolFile,"SHIFT"));
        getCell("E" + (st + 17), sheetFrom).setCellValue(get(protocolFile,"FLATNESS"));
        getCell("E" + (st + 18), sheetFrom).setCellValue(get(protocolFile,"DEVIATION9"));
        getCell("E" + (st + 19), sheetFrom).setCellValue(get(protocolFile,"DEVIATION15"));
        getCell("B" + (st + 38), sheetFrom).setCellValue(format(new Date()));
        getCell("E" + (st + 38), sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        addLimits[0] = Double.parseDouble(get(dataFile,"ADD1OUT"));
        addLimits[1] = Double.parseDouble(get(dataFile,"ADD2OUT"));
        addLimits[2] = Double.parseDouble(get(dataFile,"ADD3OUT"));
        addLimits[3] = Double.parseDouble(get(dataFile,"ADD4OUT"));
        mondLimit = Double.parseDouble(get(dataFile,"MOND"));
        monuLimit = Double.parseDouble(get(dataFile,"MONU"));
        dev9Limit = Double.parseDouble(get(dataFile,"DEVIATION9"));
        dev15Limit = Double.parseDouble(get(dataFile,"DEVIATION15"));
    }
}
