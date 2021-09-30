package test.devices.ocrm;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.format;
import static test.measure.Util.get;
import static test.measure.Util.getCell;

public class OCRM972 extends OCRM {
    public OCRM972(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OCRM/OCRM972dataGS.properties";
        } else {
            dataFile = dataDir + "OCRM/OCRM972data.properties";
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
        getCell("E20", sheetFrom).setCellValue(get(protocolFile,"1.2"));
        getCell("E21", sheetFrom).setCellValue(get(protocolFile,"1.3"));
        getCell("E22", sheetFrom).setCellValue(get(protocolFile,"1.4"));
        getCell("E23", sheetFrom).setCellValue(get(protocolFile,"1.5"));
        getCell("E24", sheetFrom).setCellValue(get(protocolFile,"1.6"));
        getCell("E25", sheetFrom).setCellValue(get(protocolFile,"1.7"));
        getCell("E26", sheetFrom).setCellValue(get(protocolFile,"1.8"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile,"1.9"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile,"2.1"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile,"2.3"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile,"2.4"));
        getCell("E31", sheetFrom).setCellValue(get(protocolFile,"2.5"));
        getCell("E32", sheetFrom).setCellValue(get(protocolFile,"2.6"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile,"2.7"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile,"2.8"));
        getCell("E35", sheetFrom).setCellValue(get(protocolFile,"2.9"));
        getCell("E36", sheetFrom).setCellValue(get(protocolFile,"3.1"));
        getCell("E37", sheetFrom).setCellValue(get(protocolFile,"3.2"));
        getCell("E38", sheetFrom).setCellValue(get(protocolFile,"3.4"));
        getCell("E39", sheetFrom).setCellValue(get(protocolFile,"3.5"));
        getCell("E40", sheetFrom).setCellValue(get(protocolFile,"3.6"));
        getCell("E41", sheetFrom).setCellValue(get(protocolFile,"3.7"));
        getCell("E42", sheetFrom).setCellValue(get(protocolFile,"3.8"));
        getCell("E43", sheetFrom).setCellValue(get(protocolFile,"3.9"));
        getCell("E44", sheetFrom).setCellValue(get(protocolFile,"4.1"));
        getCell("E45", sheetFrom).setCellValue(get(protocolFile,"4.2"));
        getCell("E46", sheetFrom).setCellValue(get(protocolFile,"4.3"));
        getCell("E47", sheetFrom).setCellValue(get(protocolFile,"4.5"));
        getCell("E48", sheetFrom).setCellValue(get(protocolFile,"4.6"));
        getCell("E49", sheetFrom).setCellValue(get(protocolFile,"4.7"));
        getCell("E50", sheetFrom).setCellValue(get(protocolFile,"4.8"));
        getCell("E51", sheetFrom).setCellValue(get(protocolFile,"4.9"));
        getCell("E52", sheetFrom).setCellValue(get(protocolFile,"5.1"));
        getCell("E53", sheetFrom).setCellValue(get(protocolFile,"5.2"));
        getCell("E54", sheetFrom).setCellValue(get(protocolFile,"5.3"));
        getCell("E55", sheetFrom).setCellValue(get(protocolFile,"5.4"));
        getCell("E56", sheetFrom).setCellValue(get(protocolFile,"5.6"));
        getCell("E57", sheetFrom).setCellValue(get(protocolFile,"5.7"));
        getCell("E58", sheetFrom).setCellValue(get(protocolFile,"5.8"));
        getCell("E59", sheetFrom).setCellValue(get(protocolFile,"5.9"));
        getCell("E60", sheetFrom).setCellValue(get(protocolFile,"6.1"));
        getCell("E61", sheetFrom).setCellValue(get(protocolFile,"6.2"));
        getCell("E62", sheetFrom).setCellValue(get(protocolFile,"6.3"));
        getCell("E63", sheetFrom).setCellValue(get(protocolFile,"6.4"));
        getCell("E65", sheetFrom).setCellValue(get(protocolFile,"6.5"));
        getCell("E66", sheetFrom).setCellValue(get(protocolFile,"6.7"));
        getCell("E67", sheetFrom).setCellValue(get(protocolFile,"6.8"));
        getCell("E68", sheetFrom).setCellValue(get(protocolFile,"6.9"));
        getCell("E69", sheetFrom).setCellValue(get(protocolFile,"7.1"));
        getCell("E70", sheetFrom).setCellValue(get(protocolFile,"7.2"));
        getCell("E71", sheetFrom).setCellValue(get(protocolFile,"7.3"));
        getCell("E72", sheetFrom).setCellValue(get(protocolFile,"7.4"));
        getCell("E73", sheetFrom).setCellValue(get(protocolFile,"7.5"));
        getCell("E74", sheetFrom).setCellValue(get(protocolFile,"7.6"));
        getCell("E75", sheetFrom).setCellValue(get(protocolFile,"7.8"));
        getCell("E76", sheetFrom).setCellValue(get(protocolFile,"7.9"));
        getCell("E77", sheetFrom).setCellValue(get(protocolFile,"8.1"));
        getCell("E78", sheetFrom).setCellValue(get(protocolFile,"8.2"));
        getCell("E79", sheetFrom).setCellValue(get(protocolFile,"8.3"));
        getCell("E80", sheetFrom).setCellValue(get(protocolFile,"8.4"));
        getCell("E81", sheetFrom).setCellValue(get(protocolFile,"8.5"));
        getCell("E82", sheetFrom).setCellValue(get(protocolFile,"8.6"));
        getCell("E83", sheetFrom).setCellValue(get(protocolFile,"8.7"));
        getCell("E84", sheetFrom).setCellValue(get(protocolFile,"8.9"));
        getCell("E85", sheetFrom).setCellValue(get(protocolFile,"9.1"));
        getCell("E86", sheetFrom).setCellValue(get(protocolFile,"9.2"));
        getCell("E87", sheetFrom).setCellValue(get(protocolFile,"9.3"));
        getCell("E88", sheetFrom).setCellValue(get(protocolFile,"9.4"));
        getCell("E89", sheetFrom).setCellValue(get(protocolFile,"9.5"));
        getCell("E90", sheetFrom).setCellValue(get(protocolFile,"9.6"));
        getCell("E91", sheetFrom).setCellValue(get(protocolFile,"9.7"));
        getCell("E92", sheetFrom).setCellValue(get(protocolFile,"9.8"));
        getCell("B182", sheetFrom).setCellValue(format(new Date()));
        getCell("E182", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile, "SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
