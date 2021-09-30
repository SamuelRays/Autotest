package test.devices.voat;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class VOAT8 extends VOA {
    public VOAT8(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "VOAT/VOAT8dataGS.properties";
            isKama = true;
        } else if (!isKama) {
            dataFile = dataDir + "VOAT/VOAT8data.properties";
        } else {
            dataFile = dataDir + "VOAT/VOAT8dataK.properties";
        }
    }

    public VOAT8(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS, isKama);
        if (isGS) {
            dataFile = dataDir + "VOAT/VOAT8dataGS.properties";
        } else if (!isKama) {
            dataFile = dataDir + "VOAT/VOAT8data.properties";
        } else {
            dataFile = dataDir + "VOAT/VOAT8dataK.properties";
        }
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
        for (int i = st; i < st + 8; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUT" + (i - st + 1)));
        }
        for (int i = st + 10; i < st + 18; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUTA" + (i - st - 9)));
        }
        for (int i = st + 21; i < st + 29; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUTF" + (i - st - 20)));
        }
        for (int i = st + 30; i < st + 38; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUT9D" + (i - st - 29)));
        }
        for (int i = st + 39; i < st + 47; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUT20D" + (i - st - 38)));
        }
        getCell("B" + (st + 75), sheetFrom).setCellValue(format(new Date()));
        getCell("E" + (st + 75), sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
