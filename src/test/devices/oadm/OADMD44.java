package test.devices.oadm;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class OADMD44 extends OADM {
    public OADMD44(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        dataFile = dataDir + "OADM/OADMD44data.properties";
    }

    public OADMD44(int slot, boolean isGS, double startCh) throws IOException {
        super(slot, isGS, startCh);
        dataFile = dataDir + "OADM/OADMD44data.properties";
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
        getCell("E21", sheetFrom).setCellValue(get(protocolFile, "CH3LN1W"));
        getCell("E22", sheetFrom).setCellValue(get(protocolFile, "CH4LN1W"));
        getCell("E23", sheetFrom).setCellValue(get(protocolFile, "LN1CH1W"));
        getCell("E24", sheetFrom).setCellValue(get(protocolFile, "LN1CH2W"));
        getCell("E25", sheetFrom).setCellValue(get(protocolFile, "LN1CH3W"));
        getCell("E26", sheetFrom).setCellValue(get(protocolFile, "LN1CH4W"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile, "CH1LN2W"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile, "CH2LN2W"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile, "CH3LN2W"));
        getCell("E31", sheetFrom).setCellValue(get(protocolFile, "CH4LN2W"));
        getCell("E32", sheetFrom).setCellValue(get(protocolFile, "LN2CH1W"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile, "LN2CH2W"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile, "LN2CH3W"));
        getCell("E35", sheetFrom).setCellValue(get(protocolFile, "LN2CH4W"));
        getCell("E38", sheetFrom).setCellValue(get(protocolFile, "CH1LN1"));
        getCell("E39", sheetFrom).setCellValue(get(protocolFile, "CH2LN1"));
        getCell("E40", sheetFrom).setCellValue(get(protocolFile, "CH3LN1"));
        getCell("E41", sheetFrom).setCellValue(get(protocolFile, "CH4LN1"));
        getCell("E42", sheetFrom).setCellValue(get(protocolFile, "UP1LN1"));
        getCell("E43", sheetFrom).setCellValue(get(protocolFile, "LN1CH1"));
        getCell("E44", sheetFrom).setCellValue(get(protocolFile, "LN1CH2"));
        getCell("E45", sheetFrom).setCellValue(get(protocolFile, "LN1CH3"));
        getCell("E46", sheetFrom).setCellValue(get(protocolFile, "LN1CH4"));
        getCell("E47", sheetFrom).setCellValue(get(protocolFile, "LN1UP1"));
        getCell("E49", sheetFrom).setCellValue(get(protocolFile, "CH1LN2"));
        getCell("E50", sheetFrom).setCellValue(get(protocolFile, "CH2LN2"));
        getCell("E51", sheetFrom).setCellValue(get(protocolFile, "CH3LN2"));
        getCell("E52", sheetFrom).setCellValue(get(protocolFile, "CH4LN2"));
        getCell("E53", sheetFrom).setCellValue(get(protocolFile, "UP2LN2"));
        getCell("E54", sheetFrom).setCellValue(get(protocolFile, "LN2CH1"));
        getCell("E55", sheetFrom).setCellValue(get(protocolFile, "LN2CH2"));
        getCell("E56", sheetFrom).setCellValue(get(protocolFile, "LN2CH3"));
        getCell("E57", sheetFrom).setCellValue(get(protocolFile, "LN2CH4"));
        getCell("E58", sheetFrom).setCellValue(get(protocolFile, "LN2UP2"));
        getCell("B91", sheetFrom).setCellValue(format(new Date()));
        getCell("F91", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
