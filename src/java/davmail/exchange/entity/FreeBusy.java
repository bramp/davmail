package davmail.exchange.entity;

import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Exchange to iCalendar Free/Busy parser.
 * Free time returns 0, Tentative returns 1, Busy returns 2, and Out of Office (OOF) returns 3
 */
public final class FreeBusy {
    final SimpleDateFormat icalParser;
    public boolean knownAttendee = true;

    public static final int INTERVAL = 15;

    static final HashMap<Character, String> FBTYPES = new HashMap<Character, String>();

    static {
        FBTYPES.put('1', "BUSY-TENTATIVE");
        FBTYPES.put('2', "BUSY");
        FBTYPES.put('3', "BUSY-UNAVAILABLE");
    }

    final HashMap<String, StringBuilder> busyMap = new HashMap<String, StringBuilder>();

    StringBuilder getBusyBuffer(char type) {
        String fbType = FBTYPES.get(Character.valueOf(type));
        StringBuilder buffer = busyMap.get(fbType);
        if (buffer == null) {
            buffer = new StringBuilder();
            busyMap.put(fbType, buffer);
        }
        return buffer;
    }

    void startBusy(char type, Calendar currentCal) {
        if (type == '4') {
            knownAttendee = false;
        } else if (type != '0') {
            StringBuilder busyBuffer = getBusyBuffer(type);
            if (busyBuffer.length() > 0) {
                busyBuffer.append(',');
            }
            busyBuffer.append(icalParser.format(currentCal.getTime()));
        }
    }

    void endBusy(char type, Calendar currentCal) {
        if (type != '0' && type != '4') {
            getBusyBuffer(type).append('/').append(icalParser.format(currentCal.getTime()));
        }
    }

    public FreeBusy(SimpleDateFormat icalParser, Date startDate, String fbdata) {
        this.icalParser = icalParser;
        if (fbdata.length() > 0) {
            Calendar currentCal = Calendar.getInstance(TimeZone.getTimeZone("UTC"));
            currentCal.setTime(startDate);

            startBusy(fbdata.charAt(0), currentCal);
            for (int i = 1; i < fbdata.length() && knownAttendee; i++) {
                currentCal.add(Calendar.MINUTE, INTERVAL);
                char previousState = fbdata.charAt(i - 1);
                char currentState = fbdata.charAt(i);
                if (previousState != currentState) {
                    endBusy(previousState, currentCal);
                    startBusy(currentState, currentCal);
                }
            }
            currentCal.add(Calendar.MINUTE, INTERVAL);
            endBusy(fbdata.charAt(fbdata.length() - 1), currentCal);
        }
    }

    /**
     * Append freebusy information to buffer.
     *
     * @param buffer String buffer
     */
    public void appendTo(StringBuilder buffer) {
        for (Map.Entry<String, StringBuilder> entry : busyMap.entrySet()) {
            buffer.append("FREEBUSY;FBTYPE=").append(entry.getKey())
                    .append(':').append(entry.getValue()).append((char) 13).append((char) 10);
        }
    }
}
