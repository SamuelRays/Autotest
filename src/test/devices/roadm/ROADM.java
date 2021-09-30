package test.devices.roadm;

import jvisa.JVisaException;
import test.devices.ActiveDevice;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.BufferedReader;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import static test.measure.Util.*;

public abstract class ROADM extends ActiveDevice {
    protected double startCh;
    protected double endCh;
    protected double shiftLimit;
    protected double flatnessLimit;
    protected double span2;
    protected String VBW2;
    protected String VBW3;
    protected double currentFlatness = 0;

    public ROADM(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isMSN = true;
    }

    protected void setChannelState(double channel, int state) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            int c = (int) Math.floor(channel);
            String on = c == channel ? "" : "e";
            setParam("WSS.Ch" + c + on + ".State.Set", "" + state);
        }
    }

    protected void setAllChState(int state) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            setParam("WSS.ChAll.State.Set", "" + state);
        }
    }

    protected void setAllChAtt(int value) throws TransformerException, ParserConfigurationException {
        if (isKama) {
            setParam("WSS.ChAll.VOA.Att.Set", "" + value);
        }
    }

    protected void allChannelTest(long time) throws IOException, TransformerException, ParserConfigurationException, InterruptedException, JVisaException {
        setAllChState(0);
        TimeUnit.MILLISECONDS.sleep(500);
        setAllChAtt(0);
        OSA.setSpan(span2);
        OSA.setVBW(VBW3);
        OSA.setTrace(5);
        double shift = 0;
        OSA.WFILON();
        for (double i = startCh; i <= endCh ; i+=1) {
            shift = channelTest(i, time, shift);
        }
        TimeUnit.MILLISECONDS.sleep(500);
        setAllChState(0);
        TimeUnit.MILLISECONDS.sleep(500);
        for (double i = startCh + 0.5; i <= endCh ; i+=1) {
            shift = channelTest(i, time, shift);
        }
        OSA.APOFF();
        TimeUnit.MILLISECONDS.sleep(500);
        setAllChState(0);
        set(protocolFile,"SHIFT", String.valueOf(round(shift, 100)));
        counter++;
    }

    protected double channelTest(double channel, long time, double shift) throws TransformerException, ParserConfigurationException, InterruptedException, JVisaException, IOException {
        while (true) {
            setChannelState(channel, 1);
            TimeUnit.MILLISECONDS.sleep(time);
            OSA.single();
            double[] results = OSA.WFILWLLVL();
            double WL = results[0];
            double lvl = results[1];
            double s = Math.abs(WL - getWL(channel));
            if (s < shiftLimit) {
                shift = s > shift ? s : shift;
               if (channel == startCh) {
                   set(protocolFile,"WLMIN", String.valueOf(round(WL, 100)));
               } else if (channel == endCh) {
                   set(protocolFile,"WLMAX", String.valueOf(round(WL, 100)));
               }
               break;
            }
        }
        return shift;
    }

    protected void lossTest(int line, String message, double down, double up, String property, BufferedReader reader) throws JVisaException, IOException, TransformerException, ParserConfigurationException {
        setAllChState(line);
        setAllChAtt(0);
        System.out.println(message);
        reader.readLine();
        if (checkThresholdsWriteResults(down, up, property, protocolFile, OSA)) {
            counter++;
        }
    }

    protected void lossFlatTest(int line, String message, double down, double up, String property, BufferedReader reader) throws JVisaException, IOException, TransformerException, ParserConfigurationException {
        setAllChState(line);
        setAllChAtt(0);
        System.out.println(message);
        reader.readLine();
        double f = getFlatThresholdsWriteResults(down, up , flatnessLimit, property, protocolFile, OSA);
        if (f < 1000) {
            currentFlatness = f > currentFlatness ? f : currentFlatness;
            counter++;
        }
    }

    protected void deviationTest(boolean isFirst, int line, int att, double limit) throws JVisaException, TransformerException, ParserConfigurationException, InterruptedException, IOException {
        setAllChState(line);
        if (isFirst) {
            setAllChAtt(0);
            OSA.setVBW(VBW2);
            OSA.takeReferenceSpectrum(0);
            TimeUnit.MILLISECONDS.sleep(5500);
        }
        setAllChAtt(att);
        TimeUnit.SECONDS.sleep(1);
        OSA.single();
        TimeUnit.SECONDS.sleep(5);
        if (checkDeviationWriteResults(att, limit, "DEVIATION" + att, protocolFile, OSA)) {
            counter++;
        }
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        startCh = Double.parseDouble(get(dataFile,"SCH"));
        endCh = Double.parseDouble(get(dataFile,"ECH"));
        flatnessLimit = Double.parseDouble(get(dataFile,"FLATNESS"));
        shiftLimit = Double.parseDouble(get(dataFile,"SHIFT"));
        span2 = Double.parseDouble(get(dataFile,"SPAN2"));
        VBW2 = get(dataFile,"VBW2");
        VBW3 = get(dataFile,"VBW3");
    }
}
