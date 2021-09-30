package test.devices.passivemux;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;
import test.devices.PassiveDevice;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class PassiveMUX extends PassiveDevice {
    protected double attLimit;
    protected double shiftLimit;
    protected boolean isPlus;

    public PassiveMUX(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
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
        while (counter != testCount) {
            String message = isPlus ? "CH" + (counter + 20) + "+" : "CH" + (counter + 20);
            System.out.println(message);
            reader.readLine();
            double channel = isPlus ? (counter + 20.5) : (counter + 20);
            if (checkWLThresholdsWriteResults(attLimit, channel, shiftLimit, "CH" + (counter + 20), "CH" + (counter + 20) + "W", protocolFile, OSA)) {
                counter++;
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
        getCell("E1", sheetFrom).setCellValue(get(protocolFile,"SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile,"PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile,"MSN"));
        for (int i = 19; i < 47; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"CH" + (i+2) + "W"));
        }
        for (int i = 48; i < 60; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"CH" + (i+1) + "W"));
        }
        for (int i = 62; i < 102; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"CH" + (i-41)));
        }
        getCell("B149", sheetFrom).setCellValue(format(new Date()));
        getCell("E149", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        attLimit = Double.parseDouble(get(dataFile, "ATT"));
        shiftLimit = Double.parseDouble(get(dataFile, "SHIFT"));
    }
}
