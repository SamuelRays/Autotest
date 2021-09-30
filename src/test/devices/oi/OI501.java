package test.devices.oi;

import jvisa.JVisaException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.*;
import java.util.Date;

import static test.measure.Util.format;
import static test.measure.Util.get;
import static test.measure.Util.getCell;

public class OI501 extends OI {
    public OI501(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        if (isGS) {
            dataFile = dataDir + "OI/OI501dataGS.properties";
        } else {
            dataFile = dataDir + "OI/OI501data.properties";
        }
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
            switch (counter) {
                case 1:
                    singleTest("CEV1", "COM IN - EVEN OUT", reader, true);
                    break;
                case 2:
                    singleTest("COD1", "COM IN - ODD OUT", reader, false);
                    break;
                case 3:
                    singleTest("ODC1", "ODD IN - COM OUT", reader, false);
                    break;
                case 4:
                    singleTest("EVC1", "EVEN IN - COM OUT", reader, true);
                    break;
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
        getCell("E19", sheetFrom).setCellValue(get(protocolFile,"CEV1W"));
        getCell("E20", sheetFrom).setCellValue(get(protocolFile,"COD1W"));
        getCell("E21", sheetFrom).setCellValue(get(protocolFile,"EVC1W"));
        getCell("E22", sheetFrom).setCellValue(get(protocolFile,"ODC1W"));
        getCell("E27", sheetFrom).setCellValue(get(protocolFile,"CEV1MIN"));
        getCell("E28", sheetFrom).setCellValue(get(protocolFile,"COD1MIN"));
        getCell("E29", sheetFrom).setCellValue(get(protocolFile,"EVC1MIN"));
        getCell("E30", sheetFrom).setCellValue(get(protocolFile,"ODC1MIN"));
        getCell("F27", sheetFrom).setCellValue(get(protocolFile,"CEV1MAX"));
        getCell("F28", sheetFrom).setCellValue(get(protocolFile,"COD1MAX"));
        getCell("F29", sheetFrom).setCellValue(get(protocolFile,"EVC1MAX"));
        getCell("F30", sheetFrom).setCellValue(get(protocolFile,"ODC1MAX"));
        getCell("E33", sheetFrom).setCellValue(get(protocolFile,"CEV1F"));
        getCell("E34", sheetFrom).setCellValue(get(protocolFile,"COD1F"));
        getCell("E35", sheetFrom).setCellValue(get(protocolFile,"EVC1F"));
        getCell("E36", sheetFrom).setCellValue(get(protocolFile,"ODC1F"));
        getCell("B50", sheetFrom).setCellValue(format(new Date()));
        getCell("F50", sheetFrom).setCellValue(engineer);
        OutputStream outputStream = new FileOutputStream(dir + get(protocolFile,"SN") + ".xlsx");
        workbookFrom.write(outputStream);
        workbookFrom.close();
    }
}
