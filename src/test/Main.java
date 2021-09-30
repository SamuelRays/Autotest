package test;

import jvisa.JVisaException;
import org.xml.sax.SAXException;
import test.devices.*;
import test.devices.activemux.OM40AVPM;
import test.devices.oadm.OADM22;
import test.devices.oadm.OADMD11;
import test.devices.oadm.OADMD44;
import test.devices.ocrm.OCRM520;
import test.devices.ocrm.OCRM972;
import test.devices.oi.OI501;
import test.devices.osc.CD;
import test.devices.osc.OSC1C;
import test.devices.osc.OSC2C;
import test.devices.passivemux.OD40A;
import test.devices.passivemux.OM40A;
import test.devices.roadm.ROADM21;
import test.devices.roadm.ROADM41;
import test.devices.roadm.ROADM91;
import test.devices.voat.VOAT2;
import test.devices.voat.VOAT4;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException, InterruptedException, SAXException, ParserConfigurationException, JVisaException, TransformerException {
        Device device = new OM40AVPM(7, false, false);
        device.test();
        device.createProtocol();
    }
}
