package com.focus.Pages;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

public class Months {

	public static void main(String[] args) {
		 Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("MMMMM-yyyy");
			String m1 = df.format(cal.getTime());
			System.out.println(m1);
			
			cal.add(Calendar.MONTH, -1);
			
			String m2 = df.format(cal.getTime());
			
			System.out.println(m2);
			cal.add(Calendar.MONTH, -1);
			
			String m3 = df.format(cal.getTime());
			System.out.println(m3);

	}

}
