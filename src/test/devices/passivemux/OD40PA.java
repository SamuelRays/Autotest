package test.devices.passivemux;

import java.io.IOException;

public class OD40PA extends PassiveMUX {
    public OD40PA(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isPlus = true;
        if (isGS) {
            dataFile = dataDir + "PassiveMUX/OD40PAdataGS.properties";
        } else {
            dataFile = dataDir + "PassiveMUX/OD40PAdata.properties";
        }
    }
}
