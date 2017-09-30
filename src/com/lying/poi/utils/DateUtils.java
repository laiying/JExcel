package com.lying.poi.utils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {
	
	public static Date formatDateString(String inDateStr, String inDateFormat) {
		Date out = null;
		SimpleDateFormat sdfin = new SimpleDateFormat(inDateFormat);
		try {
			out = sdfin.parse(inDateStr);
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if(out == null) {
			out = new Date(inDateStr);
		}
		
		return out;
	}
}
