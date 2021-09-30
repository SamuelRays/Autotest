package test.devices.passivemux;

import java.io.IOException;

public class OD40A extends PassiveMUX {
    public OD40A(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isPlus = false;
        if (isGS) {
            dataFile = dataDir + "PassiveMUX/OD40AdataGS.properties";
        } else {
            dataFile = dataDir + "PassiveMUX/OD40Adata.properties";
        }
    }
}
