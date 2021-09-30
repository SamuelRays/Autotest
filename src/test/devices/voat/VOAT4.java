package test.devices.voat;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

import static test.measure.Util.*;

public class VOAT4 extends VOA {
    public VOAT4(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "VOAT/VOAT4dataGS.properties";
            isKama = true;
        } else if (!isKama) {
            dataFile = dataDir + "VOAT/VOAT4data.properties";
        } else {
            dataFile = dataDir + "VOAT/VOAT4dataK.properties";
        }
    }

    public VOAT4(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS, isKama);
        if (isGS) {
            dataFile = dataDir + "VOAT/VOAT4dataGS.properties";
        } else if (!isKama) {
            dataFile = dataDir + "VOAT/VOAT4data.properties";
        } else {
            dataFile = dataDir + "VOAT/VOAT4dataK.properties";
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
        for (int i = st; i < st + 4; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUT" + (i - st + 1)));
        }
        for (int i = st + 6; i < st + 10; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUTA" + (i - st - 5)));
        }
        for (int i = st + 13; i < st + 17; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUTF" + (i - st - 12)));
        }
        for (int i = st + 18; i < st + 22; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUT9D" + (i - st - 17)));
        }
        for (int i = st + 23; i < st + 27; i++) {
            getCell("E" + i, sheetFrom).setCellValue(get(protocolFile,"INOUT20D" + (i - st - 22)));
        }
        getCell("B" + (st + 47), sheetFrom).setCellValue(format(new Date()));
        getCell("E" + (st + 47), sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
