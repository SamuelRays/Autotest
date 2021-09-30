package test.devices.activemux;

import java.io.IOException;

public class OM40AVPM extends ActiveMUX {
    public OM40AVPM(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isPlus = false;
        if (isGS) {
            dataFile = dataDir + "ActiveMUX/OM40AVPMdataGS.properties";
            isKama = true;
        } else if (!isKama) {
            dataFile = dataDir + "ActiveMUX/OM40AVPMdata.properties";
        } else {
            dataFile = dataDir + "ActiveMUX/OM40AVPMdataK.properties";
        }
    }

    public OM40AVPM(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS, isKama);
        isPlus = false;
        if (isGS) {
            dataFile = dataDir + "ActiveMUX/OM40AVPMdataGS.properties";
        } else if (!isKama) {
            dataFile = dataDir + "ActiveMUX/OM40AVPMdata.properties";
        } else {
            dataFile = dataDir + "ActiveMUX/OM40AVPMdataK.properties";
        }
    }
}
