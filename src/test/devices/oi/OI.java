package test.devices.oi;

import jvisa.JVisaException;
import test.devices.PassiveDevice;

import java.io.BufferedReader;
import java.io.IOException;

import static test.measure.Util.*;

public class OI extends PassiveDevice {
    protected double shiftLimit;
    protected double attLimit;
    protected double flatLimit;
    protected double currentShift;
    protected double currentMax;
    protected double currentMin;

    public OI(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isMSN = true;
    }

    protected void singleTest(String property, String message, BufferedReader reader, boolean even) throws IOException, JVisaException {
        System.out.println(message);
        reader.readLine();
        currentMax = -100;
        currentMin = 100;
        currentShift = 0;
        double ch = even ? 61 : 61.5;
        OSA.single();
        String[] WFILresults = OSA.WFIL().split(",");
        int amount = Integer.parseInt(WFILresults[0]);
        int j = 0;
        for (int i = 0; i < amount; i++) {
            double wl = Double.parseDouble(WFILresults[2 + 11 * i]);
            double lvl = -Double.parseDouble(WFILresults[5 + 11 * i]);
            double s = Math.abs(wl - getWL(ch - i + j));
            if (s > 0.2) {
                j++;
                continue;
            }
            currentShift = s > currentShift ? s : currentShift;
            currentMax = lvl > currentMax ? lvl : currentMax;
            currentMin = lvl < currentMin ? lvl : currentMin;
        }
        if (currentMax == -100) {
            System.out.println("Failed");
            return;
        }
        double flatness = currentMax - currentMin;
        System.out.println("Shift = " + currentShift);
        System.out.println("Max = " + currentMax);
        System.out.println("Min = " + currentMin);
        System.out.println("Flatness = " + flatness);
        if (currentMax <= attLimit && currentShift <= shiftLimit && flatness <= flatLimit) {
            set(protocolFile, property + "W", String.valueOf(round(currentShift, 1000)));
            set(protocolFile, property + "MIN", String.valueOf(round(currentMin, 100)));
            set(protocolFile, property + "MAX", String.valueOf(round(currentMax, 100)));
            set(protocolFile, property + "F", String.valueOf(round(flatness, 100)));
            counter++;
        } else {
            System.out.println("Failed");
        }
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        shiftLimit = Double.parseDouble(get(dataFile,"SHIFT"));
        attLimit = Double.parseDouble(get(dataFile,"ATT"));
        flatLimit = Double.parseDouble(get(dataFile,"FLATNESS"));
    }
}
