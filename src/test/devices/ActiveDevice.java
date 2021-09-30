package test.devices;

import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

import static test.measure.Util.get;

public abstract class ActiveDevice extends Device {
    protected double power;
    protected double _3v3;
    protected double _12;
    protected boolean isKama;

    public ActiveDevice(int slot, boolean isGS) throws IOException {
        super(slot, isGS);
        isKama = false;
    }

    public ActiveDevice(int slot, boolean isGS, boolean isKama) throws IOException {
        super(slot, isGS);
        this.isKama = isKama;
    }

    @Override
    protected void setPower() throws TransformerException, ParserConfigurationException, IOException, SAXException {
        setParam( "pId", pId);
        String PN = volga.CUget("Slot" + slot + "Name");
        if (!PN.endsWith("-07") && isGS) {
            PN = PN + "-07";
        }
        volga.CUset("SlotNameSet", PN);
        volga.CUset("SlotPowerSet", String.valueOf((int) power));
        volga.CUset("Slot12v0Bypass", String.valueOf(_12));
        volga.CUset("Slot3v3Bypass", String.valueOf(_3v3));
        volga.CUset("SlotWrEEPROM", String.valueOf(slot));
    }

    public void checkSum() throws IOException, TransformerException, ParserConfigurationException, SAXException {
        if (!isGS) {
            return;
        }
        String sum = getParam("MCU.DI.Hash");
        if (!sum.equals(get(protocolFile,"VPO"))) {
            System.out.println("VPO ERROR");
        } else {
            System.out.println("VPO OK");
        }
    }

    @Override
    protected void setSNs() throws IOException, TransformerException, ParserConfigurationException, SAXException {
        super.setSNs();
        setParam("SrNumber", SN);
    }

    @Override
    protected void readData() throws IOException {
        super.readData();
        power = Double.parseDouble(get(dataFile,"POWER"));
        _3v3 = Double.parseDouble(get(dataFile,"_3V3"));
        _12 = Double.parseDouble(get(dataFile,"_12"));
    }
}
