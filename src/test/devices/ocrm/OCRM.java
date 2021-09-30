package test.devices.ocrm;

import jvisa.JVisaException;
import org.xml.sax.SAXException;
import test.devices.PassiveDevice;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

import static test.measure.Util.checkThresholdsWriteResults;
import static test.measure.Util.get;

public abstract class OCRM extends PassiveDevice {
    protected double dLimit;
    protected double uLimit;
    protected int lines;

    public OCRM(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isMSN = true;
    }

    @Override
    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        super.test();
        OSA.setCenter(centerWL);
        OSA.setSpan(span);
        OSA.setVBW(VBW);
        OSA.takeReferenceSpectrum(0);
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        int out = 1;
        while (counter != testCount) {
            int in = (counter - 1) / (lines - 1) + 1;
            out++;
            out = out == in ? out + 1 : out == (lines + 1) ? 1 : out;
            System.out.println(in + " - " + out + "." + in);
            reader.readLine();
            if (checkThresholdsWriteResults(dLimit, uLimit, in + "." + out, protocolFile, OSA)) {
                counter++;
            } else {
                out--;
            }
        }
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        dLimit = Double.parseDouble(get(dataFile,"DOWN"));
        uLimit = Double.parseDouble(get(dataFile,"UP"));
        lines = Integer.parseInt(get(dataFile,"LINES"));
    }
}
