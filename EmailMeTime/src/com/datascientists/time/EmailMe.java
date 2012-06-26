package com.datascientists.time;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.util.Date;
import java.util.Timer;
import java.util.TimerTask;

public class EmailMe {
	public EmailMe() {
		
	}
	
	public void readEmails(int sinceDays) throws Exception{
		try
        {
            MailFetchIMAP gmail = new MailFetchIMAP();
            gmail.setUserPass("troy.sadkowsky@gmail.com", "xxxxxxxx");
            gmail.connect();
            gmail.openFolder("INBOX");
            gmail.readEmailsSince(sinceDays);
        }
        catch(Exception e)
        {
            e.printStackTrace();
            System.exit(-1);
        }
	}
	public void makeWeekGraph(String week,Integer iWorkSheet) throws Exception{
		Report report = new Report();
		report.removeAllValues(iWorkSheet);
		report.RunTimePeriodReport(week,iWorkSheet);
	}
	public static void main(String[] argv) {
		System.out.println("Running EmailMeTime version 0.01");
		Timer timer = new Timer();
		EmailMeTimer emt = new EmailMeTimer();
		timer.scheduleAtFixedRate(emt, new Date(), (3*60*60*1000));//every three hours
	}
}
class EmailMeTimer extends TimerTask {
	@Override
	public void run() {
		EmailMe emailMe = new EmailMe();
		try {
			String[] inputs = new String[3];
			FileInputStream fstream = new FileInputStream("/home/occideas/Desktop/emailme.config");
			// Get the object of DataInputStream
			DataInputStream in = new DataInputStream(fstream);
			BufferedReader br = new BufferedReader(new InputStreamReader(in));
			String strLine;
			int i=0;
			//Read File Line By Line
			while ((strLine = br.readLine()) != null)   {
			  // Print the content on the console
			  inputs[i]=strLine;
			  i++;
			}
			//Close the input stream
			in.close();
			emailMe.readEmails(Integer.valueOf(inputs[0]));
			System.out.println("Processing week "+inputs[1]);
			System.out.println("Wrting to sheet "+inputs[2]);
			emailMe.makeWeekGraph(inputs[1],Integer.valueOf(inputs[2]));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("Waiting...");

	}

}
