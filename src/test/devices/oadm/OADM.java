package test.devices.oadm;

import jvisa.JVisaException;
import org.xml.sax.SAXException;
import test.devices.PassiveDevice;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

import static test.measure.Util.*;

public abstract class OADM extends PassiveDevice {
    protected double shiftLimit;
    protected double cLimit;
    protected double uLimit;
    protected double startCh;
    protected int cAmount;
    protected int lAmount;

    public OADM(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        startCh = 20;
        isMSN = true;
    }

    public OADM(int slot, boolean isGS, double startCh) throws IOException {
        super(slot, isGS);
        this.startCh = startCh;
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
        for (int i = 1; i < lAmount + 1; i++) {
            for (int k = 0; k < 2; k++) {
                boolean forward = k == 0;
                for (int j = 1; j < cAmount + 2; j++) {
                    if (j < cAmount + 1) {
                        if (!Wtest(startCh - 1 + j, j, i, reader, forward)) {
                            j--;
                        }
                    } else {
                        if (!Utest(i, reader, forward)) {
                            j--;
                        }
                    }
                }
            }
        }
        OSA.APOFF();
    }

    protected boolean Wtest(double ch, int c, int line, BufferedReader reader, boolean forward) throws IOException, JVisaException {
        String message = forward ? "CH " + (int) ch + " IN - LINE" + line + " OUT" : "LINE" + line + " IN - CH " + (int) ch + " OUT";
        String property = forward ? "CH" + c + "LN" + line : "LN" + line + "CH" + c;
        System.out.println(message);
        reader.readLine();
        return checkWLThresholdsWriteResults(cLimit, ch, shiftLimit, property, property + "W", protocolFile, OSA);
    }

    protected boolean Utest(int line, BufferedReader reader, boolean forward) throws IOException, JVisaException {
        String message = forward ? "UPG " + line + " IN - LINE" + line + " OUT" : "LINE" + line + " IN - UPG " + line + " OUT";
        String property = forward ? "UP" + line + "LN" + line : "LN" + line + "UP" + line;
        System.out.println(message);
        reader.readLine();
        return checkThresholdWriteResults(0.1, uLimit, property, protocolFile, OSA);
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        shiftLimit = Double.parseDouble(get(dataFile,"SHIFT"));
        cLimit = Double.parseDouble(get(dataFile,"C"));
        uLimit = Double.parseDouble(get(dataFile,"U"));
        cAmount = Integer.parseInt(get(dataFile,"CAMOUNT"));
        lAmount = Integer.parseInt(get(dataFile,"LAMOUNT"));
        if (!isGS) {
            if (lAmount == 1) {
                pId = pId + (int) startCh + "-01";
            } else {
                pId = pId + (int) startCh + "/C" + (int) startCh + "-01";
            }
        }
    }
}
