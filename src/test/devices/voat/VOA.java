package test.devices.voat;

import jvisa.JVisaException;
import org.xml.sax.SAXException;
import test.devices.ActiveDevice;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.concurrent.TimeUnit;

import static test.measure.Util.*;

public abstract class VOA extends ActiveDevice {
    protected double inoutLimit;
    protected double inoutALimit;
    protected double flatnessLimit;
    protected double dev9Limit;
    protected double dev20Limit;
    protected int VOAAmount;

    public VOA(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isMSN = true;
    }

    public VOA(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS, isKama);
        isMSN = true;
    }

    @Override
    public void test() throws InterruptedException, TransformerException, SAXException, JVisaException, ParserConfigurationException, IOException {
        super.test();
        OSA.setCenter(centerWL);
        OSA.setSpan(span);
        OSA.setVBW(VBW);
        OSA.takeReferenceSpectrum(0);
        VOAON();
        setAllVOAAtt(0);
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        while (counter != testCount) {
            System.out.printf("IN %d - OUT %d\n", counter, counter);
            reader.readLine();
            VOATest(counter);
        }
        VOAOFF();
        volga.logout();
        OSA.logout();
    }

    protected void VOATest(int VOA) throws IOException, JVisaException, TransformerException, ParserConfigurationException, InterruptedException, SAXException {
        if (checkFlatThresholdsWriteResults(0.1, inoutLimit, flatnessLimit, "INOUT" + VOA, "INOUTF" + VOA, protocolFile, OSA)) {
            power = Math.abs(OSA.getPower() - getPower(VOA));
            System.out.println(power);
            if (power > inoutALimit) {
                System.out.println("Failed");
            }
            set(protocolFile, "INOUTA" + VOA, String.valueOf(round(power, 10)));
            OSA.takeReferenceSpectrum(3);
            TimeUnit.SECONDS.sleep(4);
            setVOAAtt(VOA, 9);
            deviationTest(VOA, 9, "INOUT9D", dev9Limit);
            deviationTest(VOA, 20, "INOUT20D", dev20Limit);
            OSA.setTrace(2);
            counter++;
        }
    }

    protected void deviationTest(int VOA, double att, String property, double limit) throws JVisaException, TransformerException, ParserConfigurationException, InterruptedException, IOException {
        setVOAAtt(VOA, att);
        TimeUnit.SECONDS.sleep(2);
        OSA.single();
        double dev = OSA.deviation(att);
        System.out.println(dev);
        if (dev > limit) {
            System.out.println("Failed");
        }
        set(protocolFile,property + VOA, String.valueOf(round(dev, 10)));
    }

    protected void VOAON() throws TransformerException, ParserConfigurationException, InterruptedException {
        setAllVOAState(1);
        if (!isKama) {
            TimeUnit.SECONDS.sleep(1);
        }
        setAllVOAState(2);
    }

    protected void VOAOFF() throws TransformerException, ParserConfigurationException, InterruptedException {
        setAllVOAState(1);
        if (!isKama) {
            TimeUnit.SECONDS.sleep(1);
        }
        setAllVOAState(0);
    }

    protected void setAllVOAState(int state) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            for (int i = 1; i < VOAAmount + 1; i++) {
                setVOAState(i, state);
            }
        } else {
            setParam("VOAAll.Port.State", "" + state);
        }
    }

    protected void setVOAState(int VOA, int state) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            setParam("VOA" + VOA + ".AdmState.Set", "" + state);
        } else {
            setParam("VOA" + VOA + "Out.Port.State", "" + state);
        }
    }

    protected double getPower(int VOA) throws ParserConfigurationException, IOException, SAXException, TransformerException {
        double result = 0;
        if (isKama) {
            result = Double.parseDouble(getParam("VOA" + VOA + ".Out.Power"));
        } else {
            result = Double.parseDouble(getParam("VOAPort" + VOA + "Pout"));
        }
        return result;
    }

    protected void setAllVOAAtt(double att) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            for (int i = 1; i < VOAAmount + 1; i++) {
                setVOAAtt(i, att);
            }
        } else {
            setParam("VOAAll.Att.Set", "" + att);
        }
    }

    protected void setVOAAtt(int VOA, double att) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            setParam("VOA" + VOA + ".Att.Set", "" + att);
        } else {
            setParam("VOAPort" + VOA + "AttSet", "" + att);
        }
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        inoutLimit = Double.parseDouble(get(dataFile,"INOUT"));
        inoutALimit = Double.parseDouble(get(dataFile,"INOUTA"));
        flatnessLimit = Double.parseDouble(get(dataFile,"FLATNESS"));
        dev9Limit = Double.parseDouble(get(dataFile,"DEVIATION9"));
        dev20Limit = Double.parseDouble(get(dataFile,"DEVIATION20"));
        VOAAmount = Integer.parseInt(get(dataFile,"VOA"));
    }
}
