package test.devices.osc;

import jvisa.JVisaException;
import test.devices.PassiveDevice;

import java.io.BufferedReader;
import java.io.IOException;

import static test.measure.Util.*;

public abstract class OSC extends PassiveDevice {
    protected double OSCLimit;
    protected double upLimit;
    protected double mon1dLimit;
    protected double mon1uLimit;
    protected double mon2dLimit;
    protected double mon2uLimit;
    protected double span2;
    protected double centerWL2;
    protected double SFPPeak;

    public OSC(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isMSN = true;
    }

    protected void OSCCTest(double down, double up, String message, String property, BufferedReader reader) throws IOException, JVisaException {
        System.out.println(message);
        reader.readLine();
        if (checkShiftThresholdWriteResults(SFPPeak, down, up, property, protocolFile, OSA)) {
            counter++;
        }
    }

    protected void upCTest(double down, double up, String message, String property, BufferedReader reader) throws IOException, JVisaException {
        System.out.println(message);
        reader.readLine();
        if (checkThresholdsWriteResults(down, up, property, protocolFile, OSA)) {
            counter++;
        }
    }

    protected void OSCDTest(double down, double ch, double shiftLimit, String message, String property, BufferedReader reader) throws IOException, JVisaException {
        System.out.println(message);
        reader.readLine();
        if (checkWLThresholdsWriteResults(down, ch, shiftLimit, property, property + "W", protocolFile, OSA)) {
            counter++;
        }
    }

    protected void upDTest(double up, String message, String property, BufferedReader reader) throws IOException, JVisaException {
        System.out.println(message);
        reader.readLine();
        if (checkThresholdWriteResults(0.1, up, property, protocolFile, OSA)) {
            counter++;
        }
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        OSCLimit = Double.parseDouble(get(dataFile,"OSC"));
        upLimit = Double.parseDouble(get(dataFile,"UP"));
        mon1dLimit = Double.parseDouble(get(dataFile,"MON1D"));
        mon1uLimit = Double.parseDouble(get(dataFile,"MON1U"));
        mon2dLimit = Double.parseDouble(get(dataFile,"MON2D"));
        mon2uLimit = Double.parseDouble(get(dataFile,"MON2U"));
        span2 = Double.parseDouble(get(dataFile,"SPAN2"));
        centerWL2 = Double.parseDouble(get(dataFile,"CENTER2"));
    }
}
