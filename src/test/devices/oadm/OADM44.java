package test.devices.oadm;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class OADM44 extends OADM {
    public OADM44(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        dataFile = dataDir + "OADM/OADM44data.properties";
    }

    public OADM44(int slot, boolean isGS, double startCh) throws IOException {
        super(slot, isGS, startCh);
        dataFile = dataDir + "OADM/OADM44data.properties";
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        getCell("H1", sheetFrom).setCellValue((int) startCh);
        getCell("E1", sheetFrom).setCellValue(get(protocolFile, "SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile, "PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile, "MSN"));
        for (int i = 1; i < 5; i++) {
            getCell("E" + (i + 18), sheetFrom).setCellValue(get(protocolFile, "CH" + i + "LN1W"));
        }
        for (int i = 1; i < 5; i++) {
            getCell("E" + (i + 22), sheetFrom).setCellValue(get(protocolFile, "LN1" + "CH" + i + "W"));
        }
        for (int i = 1; i < 5; i++) {
            getCell("E" + (i + 28), sheetFrom).setCellValue(get(protocolFile, "CH" + i + "LN1"));
        }
        getCell("E33", sheetFrom).setCellValue(get(protocolFile, "UP1LN1"));
        for (int i = 1; i < 5; i++) {
            getCell("E" + (i + 33), sheetFrom).setCellValue(get(protocolFile, "LN1CH" + i));
        }
        getCell("E38", sheetFrom).setCellValue(get(protocolFile, "LN1UP1"));
        getCell("B58", sheetFrom).setCellValue(format(new Date()));
        getCell("E58", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
