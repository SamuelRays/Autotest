package test.devices.oadm;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class OADMD11 extends OADM {
    public OADMD11(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OADM/OADMD11dataGS.properties";
        } else {
            dataFile = dataDir + "OADM/OADMD11data.properties";
        }
    }

    public OADMD11(int slot, boolean isGS, double startCh) throws IOException {
        super(slot, isGS, startCh);
        if (isGS) {
            dataFile = dataDir + "OADM/OADMD11dataGS.properties";
        } else {
            dataFile = dataDir + "OADM/OADMD11data.properties";
        }
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        if (!isGS) {
            getCell("H1", sheetFrom).setCellValue((int) startCh);
        }
        getCell("E1", sheetFrom).setCellValue(get(protocolFile, "SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile, "PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile, "MSN"));
        getCell("E19", sheetFrom).setCellValue(get(protocolFile, "CH1LN1W"));
        getCell("E20", sheetFrom).setCellValue(get(protocolFile, "LN1CH1W"));
        getCell("E22", sheetFrom).setCellValue(get(protocolFile, "CH1LN2W"));
        getCell("E23", sheetFrom).setCellValue(get(protocolFile, "LN2CH1W"));
        getCell("E26", sheetFrom).setCellValue(get(protocolFile, "CH1LN1"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile, "UP1LN1"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile, "LN1CH1"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile, "LN1UP1"));
        getCell("E31", sheetFrom).setCellValue(get(protocolFile, "CH1LN2"));
        getCell("E32", sheetFrom).setCellValue(get(protocolFile, "UP2LN2"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile, "LN2CH1"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile, "LN2UP2"));
        getCell("B49", sheetFrom).setCellValue(format(new Date()));
        getCell("E49", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
