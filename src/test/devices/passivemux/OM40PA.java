package test.devices.passivemux;

import java.io.IOException;

public class OM40PA extends PassiveMUX {
    public OM40PA(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isPlus = true;
        if (isGS) {
            dataFile = dataDir + "PassiveMUX/OM40PAdataGS.properties";
        } else {
            dataFile = dataDir + "PassiveMUX/OM40PAdata.properties";
        }
    }
}
