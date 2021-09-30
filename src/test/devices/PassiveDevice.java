package test.devices;

import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import static test.measure.Util.get;

public abstract class PassiveDevice extends Device {
    private String className;

    public PassiveDevice(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
    }

    @Override
    protected void setPower() throws TransformerException, ParserConfigurationException, IOException, SAXException, InterruptedException {
        String PN = volga.CUget("Slot" + slot + "Name");
        if (!PN.endsWith("-07") && isGS) {
            PN = PN + "-07";
        }
        volga.CUset("SlotNameSet", PN);
        volga.CUset("SlotPowerSet", String.valueOf(0));
        volga.CUset("Slot12v0Bypass", String.valueOf(50));
        volga.CUset("Slot3v3Bypass", String.valueOf(50));
        volga.CUset("SlotWrEEPROM", String.valueOf(slot));
        TimeUnit.SECONDS.sleep(1);
    }

    @Override
    protected void setSNs() throws IOException, TransformerException, ParserConfigurationException, SAXException {
        super.setSNs();
        volga.CUset("Slot" + slot, String.format("class=\"%s\",pId=\"%s\",sernum=\"%s\",dest=\"\",desc=\"\"", className, pId, SN));
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        className = get(dataFile,"CLASS");
    }
}
