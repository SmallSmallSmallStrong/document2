package model01;

public class Util {

    public static String docxname(String s) {
        String ret = s.substring(0, 1) + " (" + s.substring(1, s.length() - 1) + ")";
        return ret;
    }
}
