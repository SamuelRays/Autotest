package test.devices.activemux;

import java.io.IOException;

public class OM40PAVPM extends ActiveMUX {
    public OM40PAVPM(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isPlus = true;
        if (isGS) {
            dataFile = dataDir + "ActiveMUX/OM40PAVPMdataGS.properties";
            isKama = true;
        } else if (!isKama) {
            dataFile = dataDir + "ActiveMUX/OM40PAVPMdata.properties";
        } else {
            dataFile = dataDir + "ActiveMUX/OM40PAVPMdataK.properties";
        }
    }

    public OM40PAVPM(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS, isKama);
        isPlus = true;
        if (isGS) {
            dataFile = dataDir + "ActiveMUX/OM40PAVPMdataGS.properties";
        } else if (!isKama) {
            dataFile = dataDir + "ActiveMUX/OM40PAVPMdata.properties";
        } else {
            dataFile = dataDir + "ActiveMUX/OM40PAVPMdataK.properties";
        }
    }
}
