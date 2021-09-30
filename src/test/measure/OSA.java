package test.measure;

import jvisa.JVisa;
import jvisa.JVisaException;
import jvisa.JVisaReturnString;

import java.util.concurrent.TimeUnit;

public class OSA {
    private static String[] traces = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"};
    private String IP = "TCPIP::192.168.1.207::INSTR";
    private String VBW;
    private double span;
    private double center;
    private String trace;
    private JVisa visa;
    private JVisaReturnString returnString = new JVisaReturnString();

    public OSA(String IP) {
        this.IP = "TCPIP::" + IP + "::INSTR";
    }

    public OSA() {
    }

    public void connect() throws JVisaException {
        visa = new JVisa();
        visa.openDefaultResourceManager();
        visa.openInstrument(IP);
        visa.setTimeout(300000);
        visa.write("VBW?");
        visa.read(returnString);
        VBW = returnString.returnString;
        visa.write("CNT?");
        visa.read(returnString);
        center = Double.parseDouble(returnString.returnString);
        visa.write("SPN?");
        visa.read(returnString);
        span = Double.parseDouble(returnString.returnString);
        visa.write("TSL?");
        visa.read(returnString);
        trace = returnString.returnString;
    }

    public void logout() {
        visa.closeInstrument();
        visa.closeResourceManager();
    }

    public void takeReferenceSpectrum(int trace1) throws JVisaException {
        int trace2 = trace1 + 1;
        int trace3 = trace1 + 2;
        visa.write("TSL " + traces[trace1]);
        visa.write("TMD " + traces[trace1] + ", ON");
        visa.write("TTP " + traces[trace1] + ", WRITE");
        visa.write("SSI; *WAI");
        visa.write("TTP " + traces[trace1] + ", FIX");
        visa.write("TTP " + traces[trace2] + ", WRITE");
        visa.write("TTP " + traces[trace3] + ", CALC");
        visa.write("FML " + traces[trace3] + "," + traces[trace2] + ",-," + traces[trace1]);
        visa.write("TSL " + traces[trace3]);
        visa.write("TMD " + traces[trace3] + ", ON");
    }

    public void single() throws JVisaException {
        visa.write("SSI; *WAI");
    }

    public double getPower() throws JVisaException, InterruptedException {
        visa.write("PWR 1550; *WAI");
        TimeUnit.SECONDS.sleep(8);
        visa.write("PWRR?");
        visa.read(returnString, 20480);
        double power = Double.parseDouble(returnString.returnString);
        visa.write("SPC");
        return power;
    }

    public void WFILON() throws JVisaException {
        visa.write("AP WFIL");
    }

    public String WFIL() throws JVisaException {
        visa.write("APR? WFIL");
        visa.read(returnString, 20480);
        return returnString.returnString;
    }

    public void APOFF() throws JVisaException {
        visa.write("AP OFF");
    }

    public double[] WFILWLLVL() throws JVisaException {
        String r = WFIL();
        double[] result = new double[2];
        result[0] = Double.parseDouble(r.split(",")[2]);
        result[1] = Double.parseDouble(r.split(",")[5]);
        return result;
    }

    public double peakSearch() throws JVisaException {
        visa.write("PKS PEAK; *WAI; TMK?");
        visa.read(returnString);
        String result = returnString.returnString.split(",")[1];
        return Double.parseDouble(result.substring(0, result.length() - 2));
    }

    public double dipSearch() throws JVisaException {
        visa.write("DPS DIP; *WAI; TMK?");
        visa.read(returnString);
        String result = returnString.returnString.split(",")[1];
        return Double.parseDouble(result.substring(0, result.length() - 2));
    }

    public double mean() throws JVisaException {
        return (peakSearch() + dipSearch()) / 2;
    }

    public double flatness() throws JVisaException {
        return peakSearch() - dipSearch();
    }

    public double deviation(double value) throws JVisaException {
        return Math.max(-dipSearch() - value, peakSearch() + value);
    }

    public String getVBW() {
        return VBW;
    }

    public void setVBW(String VBW) throws JVisaException {
        visa.write("VBW " + VBW);
        this.VBW = VBW;
    }

    public double getSpan() {
        return span;
    }

    public void setSpan(double span) throws JVisaException {
        visa.write("SPN " + span);
        this.span = span;
    }

    public double getCenter() {
        return center;
    }

    public void setCenter(double center) throws JVisaException {
        visa.write("CNT " + center);
        this.center = center;
    }

    public String getTrace() {
        return trace;
    }

    public void setTrace(int trace) throws JVisaException {
        visa.write("TSL " + traces[trace]);
        visa.write("TMD " + traces[trace] + ", ON");
        this.trace = traces[trace];
    }
}
