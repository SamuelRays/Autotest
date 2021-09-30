package test.devices.oadm;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class OADM22 extends OADM {
    public OADM22(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        dataFile = dataDir + "OADM/OADM22data.properties";
    }

    public OADM22(int slot, boolean isGS, double startCh) throws IOException {
        super(slot, isGS, startCh);
        dataFile = dataDir + "OADM/OADM22data.properties";
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
        getCell("E19", sheetFrom).setCellValue(get(protocolFile, "CH1LN1W"));
        getCell("E20", sheetFrom).setCellValue(get(protocolFile, "CH2LN1W"));
        getCell("E21", sheetFrom).setCellValue(get(protocolFile, "LN1CH1W"));
        getCell("E22", sheetFrom).setCellValue(get(protocolFile, "LN1CH2W"));
        getCell("E25", sheetFrom).setCellValue(get(protocolFile, "CH1LN1"));
        getCell("E26", sheetFrom).setCellValue(get(protocolFile, "CH2LN1"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile, "UP1LN1"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile, "LN1CH1"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile, "LN1CH2"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile, "LN1UP1"));
        getCell("B46", sheetFrom).setCellValue(format(new Date()));
        getCell("E46", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
