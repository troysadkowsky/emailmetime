package com.datascientists.time;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.Writer;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.URLName;
import javax.mail.search.AndTerm;
import javax.mail.search.ComparisonTerm;
import javax.mail.search.FromStringTerm;
import javax.mail.search.OrTerm;
import javax.mail.search.ReceivedDateTerm;
import javax.mail.search.SearchTerm;
import javax.mail.search.SubjectTerm;

import com.google.gdata.client.DocumentQuery;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.ListEntry;
import com.google.gdata.data.spreadsheet.ListFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.sun.mail.imap.IMAPFolder;
import com.sun.mail.imap.IMAPStore;

public class MailFetchIMAP
{
    private Session session;
    private IMAPStore store;
    private String username;
    private String password;
    private IMAPFolder folder;
    public static String numberOfFiles = null;
    public static int toCheck = 0;
    public static Writer output = null;
    URLName url;
    public static String receiving_attachments="C:\\download";
    private Calendar lastTimeStamp;

    public MailFetchIMAP()
    {
        session = null;
        store = null;
    }

    public void setUserPass(String username, String password)
    {
        this.username = username;
        this.password = password;
    }

    public void connect()
    throws Exception
    {
        Properties props = new Properties();
        props.setProperty("mail.store.protocol", "imaps");
		props.setProperty("mail.imaps.host", "imap.gmail.com");
		props.setProperty("mail.imaps.port", "993");
		props.setProperty("mail.imaps.partialfetch", "false");

		session = javax.mail.Session.getDefaultInstance(props, null);
		store = (IMAPStore) session.getStore("imaps");

		store.connect(username, password);
    }

    public void openFolder(String folderName)
    throws Exception
    {
        folder = (IMAPFolder)store.getFolder(folderName);
        if(folder == null)
            throw new Exception("Invalid folder");
        try
        {
            folder.open(Folder.READ_ONLY);             
        }
        catch(Exception ex)
        {
            System.out.println((new StringBuilder("Folder Opening Exception..")).append(ex).toString());
        }
    }

    public void closeFolder()
    throws Exception
    {
        folder.close(false);
    }

    public int getMessageCount()
    throws Exception
    {
        return folder.getMessageCount();
    }

    public int getNewMessageCount()
    throws Exception
    {
        return folder.getNewMessageCount();
    }

    public void disconnect()
    throws Exception
    {
        store.close();
    }

    public void readEmailsSince(int sinceDays)
    throws Exception
    {
    	javax.mail.Folder[] folders = store.getDefaultFolder().list("EmailMe");
        for (javax.mail.Folder folder : folders) {
            if ((folder.getType() & javax.mail.Folder.HOLDS_MESSAGES) != 0) {
                System.out.println("Reading emails since "+sinceDays+" days ago in "+folder.getFullName());
                System.out.println("Total count: " + folder.getMessageCount());
                SearchTerm term = null;
                term = new SubjectTerm("Project"); 
                FromStringTerm fromTerm = new FromStringTerm("troy.sadkowsky@me.com");
                term = new AndTerm(term, fromTerm);
                Calendar date = Calendar.getInstance();
                ReceivedDateTerm dateTerm = new ReceivedDateTerm(ComparisonTerm.EQ, date.getTime());      
                term = new AndTerm(term, dateTerm);
                for(int i=0;i<sinceDays;i++){
                	date.add(Calendar.DATE, -1);
                    dateTerm = new ReceivedDateTerm(ComparisonTerm.EQ, date.getTime());      
                    term = new OrTerm(term, dateTerm);
                }                        
                folder.open(Folder.READ_ONLY);
                Message msgs[] = folder.search(term);
                System.out.println("Found count: " + msgs.length);
                this.writeToGoogleDoc(msgs); 
                folder.close(true);
            }
        }
    }



    public static int saveFile(File saveFile, Part part) throws Exception {

        BufferedOutputStream bos = new BufferedOutputStream( new FileOutputStream(saveFile) );

        byte[] buff = new byte[2048];
        InputStream is = part.getInputStream();
        int ret = 0, count = 0;
        while( (ret = is.read(buff)) > 0 ){
            bos.write(buff, 0, ret);
            count += ret;
        }
        bos.close();
        is.close();
        return count;
    }

    public void writeToGoogleDoc(Message msgs[]) throws Exception
    {
    	URL metafeedUrl = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");
		SpreadsheetService service = new SpreadsheetService("EmailMeTime");
		service.setUserCredentials("troy.sadkowsky@gmail.com", "p34s0up");
		DocumentQuery query = new DocumentQuery(metafeedUrl);
		query.setTitleQuery("EmailMeTime");
		query.setTitleExact(true);
		query.setMaxResults(1);

		SpreadsheetFeed feed = service.getFeed(query, SpreadsheetFeed.class);
		List<SpreadsheetEntry> spreadsheets = feed.getEntries();
		for (int j = 0; j < spreadsheets.size(); j++) {
			SpreadsheetEntry entry = spreadsheets.get(j);
			
			List<WorksheetEntry> worksheets = entry.getWorksheets();
			if (worksheets.size() > 1) {
				WorksheetEntry worksheetEntry = worksheets.get(0);
				URL listFeedUrl = worksheetEntry.getListFeedUrl();

				ListFeed wfeed2 = service.getFeed(listFeedUrl, ListFeed.class);

				String strDateOfLastEntry = "";
				for (ListEntry entry1 : wfeed2.getEntries()) {
					// System.out.println(entry1.getTitle().getPlainText());
					strDateOfLastEntry = entry1.getCustomElements().getValue(
							"Date");
				}
				DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
				Date dateOfLastEntry = (Date)df.parse(strDateOfLastEntry); 

				for (int i = 0; i < msgs.length; i++) {

					String subject = msgs[i].getSubject();
					
					Date date = msgs[i].getSentDate();
					
					String formattedDate = df.format(date);
					
					DateFormat dfw = new SimpleDateFormat("yyyy'W'ww");
					String week = dfw.format(date);
					
					Object content = msgs[i].getContent();
					String body = "";
					if (content instanceof String) {

						body = (String) content;
						
					}
					Calendar cal = Calendar.getInstance();
					cal.setTime(date);

					Integer idayOfWeek = cal.get(Calendar.DAY_OF_WEEK);
					String[] daysOfTheWeek = { "", "Sunday", "Monday",
							"Tuesday", "Wednesday", "Thursday", "Friday",
							"Saturday" };
					String dayOfTheWeek = daysOfTheWeek[idayOfWeek];
					String firstLine = body.split("\n")[0];
					int iLastSpace = firstLine.trim().lastIndexOf(" ");
					String duration = "0.5";
					try {
						duration = firstLine.substring(iLastSpace);
					} catch (Exception e) {
						System.out.println("No time using default of 0.5");
					}
					
					worksheetEntry = worksheets.get(0);
					listFeedUrl = worksheetEntry.getListFeedUrl();

					ListEntry newEntry = new ListEntry();
					newEntry.getCustomElements().setValueLocal("Project",
							subject);
					newEntry.getCustomElements().setValueLocal("Date",
							formattedDate);
					newEntry.getCustomElements().setValueLocal("DayOfWeek",
							dayOfTheWeek);
					newEntry.getCustomElements().setValueLocal("Week", week);
					newEntry.getCustomElements().setValueLocal("Body", body);
					newEntry.getCustomElements().setValueLocal("FirstLine",
							firstLine);
					newEntry.getCustomElements().setValueLocal("Duration",
							duration);

					float hoursBetween = 0;
					if (lastTimeStamp != null) {
						Calendar dateTemp = (Calendar) lastTimeStamp.clone();
						long minutesBetween = 0;
						while (dateTemp.before(cal)) {
							dateTemp.add(Calendar.MINUTE, 1);
							minutesBetween++;
						}

						hoursBetween = ((float) minutesBetween / (float) 60);
					}
					String timeSinceLast = String.valueOf(hoursBetween);
					newEntry.getCustomElements().setValueLocal(
							"TimeSinceLastEntry", timeSinceLast);
					
					//String a = cal.toString();
					//String b = dateOfLastEntry.toString();
					Calendar calOfLastEntry = Calendar.getInstance();
					calOfLastEntry.setTime(dateOfLastEntry);
					if (cal.after(calOfLastEntry)) {
						service.insert(listFeedUrl, newEntry);
						System.out.println("New Entry Found:"+subject+duration);
						System.out.print(body);
					}
					this.lastTimeStamp = cal;

				}
			}
		}
	}
    
}
