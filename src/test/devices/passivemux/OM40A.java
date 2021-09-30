package test.devices.passivemux;

import java.io.IOException;

public class OM40A extends PassiveMUX {
    public OM40A(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isPlus = false;
        if (isGS) {
            dataFile = dataDir + "PassiveMUX/OM40AdataGS.properties";
        } else {
            dataFile = dataDir + "PassiveMUX/OM40Adata.properties";
        }
    }
}
