package test.devices.ocrm;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.format;
import static test.measure.Util.get;
import static test.measure.Util.getCell;

public class OCRM520 extends OCRM {
    public OCRM520(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OCRM/OCRM520dataGS.properties";
        } else {
            dataFile = dataDir + "OCRM/OCRM520data.properties";
        }
    }

    @Override
    public void createProtocol() throws IOException {
        super.createProtocol();
        InputStream inputStreamFrom = new FileInputStream(sampleFile);
        Workbook workbookFrom = new XSSFWorkbook(inputStreamFrom);
        workbookFrom.setForceFormulaRecalculation(true);
        Sheet sheetFrom = workbookFrom.getSheetAt(0);
        getCell("E1", sheetFrom).setCellValue(get(protocolFile, "SN"));
        getCell("C4", sheetFrom).setCellValue(get(protocolFile, "PN"));
        getCell("C5", sheetFrom).setCellValue(get(protocolFile, "MSN"));
        getCell("E20", sheetFrom).setCellValue(get(protocolFile, "1.2"));
        getCell("E21", sheetFrom).setCellValue(get(protocolFile, "1.3"));
        getCell("E22", sheetFrom).setCellValue(get(protocolFile, "1.4"));
        getCell("E23", sheetFrom).setCellValue(get(protocolFile, "1.5"));
        getCell("E24", sheetFrom).setCellValue(get(protocolFile, "2.1"));
        getCell("E25", sheetFrom).setCellValue(get(protocolFile, "2.3"));
        getCell("E26", sheetFrom).setCellValue(get(protocolFile, "2.4"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile, "2.5"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile, "3.1"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile, "3.2"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile, "3.4"));
        getCell("E31", sheetFrom).setCellValue(get(protocolFile, "3.5"));
        getCell("E32", sheetFrom).setCellValue(get(protocolFile, "4.1"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile, "4.2"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile, "4.3"));
        getCell("E35", sheetFrom).setCellValue(get(protocolFile, "4.5"));
        getCell("E36", sheetFrom).setCellValue(get(protocolFile, "5.1"));
        getCell("E37", sheetFrom).setCellValue(get(protocolFile, "5.2"));
        getCell("E38", sheetFrom).setCellValue(get(protocolFile, "5.3"));
        getCell("E39", sheetFrom).setCellValue(get(protocolFile, "5.4"));
        getCell("B72", sheetFrom).setCellValue(format(new Date()));
        getCell("E72", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile, "SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
