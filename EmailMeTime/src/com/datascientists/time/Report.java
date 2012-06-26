package com.datascientists.time;

import java.net.URL;
import java.util.List;

import com.google.gdata.client.DocumentQuery;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.ListEntry;
import com.google.gdata.data.spreadsheet.ListFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;

public class Report {

	public Report(){
		
	}
	public void removeAllValues(Integer iWorkSheet) throws Exception {
		URL metafeedUrl = new URL(
				"https://spreadsheets.google.com/feeds/spreadsheets/private/full");
		SpreadsheetService service = new SpreadsheetService("EmailMeTime");
		service.setUserCredentials("troy.sadkowsky@gmail.com", "p34s0up");
		DocumentQuery query = new DocumentQuery(metafeedUrl);
		query.setTitleQuery("EmailMeTime");
		query.setTitleExact(true);
		query.setMaxResults(1);

		String[] daysOfTheWeek = { "", "Sunday", "Monday", "Tuesday",
				"Wednesday", "Thursday", "Friday", "Saturday" };

		SpreadsheetFeed feed = service.getFeed(query, SpreadsheetFeed.class);
		List<SpreadsheetEntry> spreadsheets = feed.getEntries();
		for (int i = 0; i < spreadsheets.size(); i++) {
			SpreadsheetEntry entry = spreadsheets.get(i);
			// System.out.println(entry.getTitle().getPlainText());
			List<WorksheetEntry> worksheets = entry.getWorksheets();
			//for (int j = 0; j < worksheets.size(); j++) {
				WorksheetEntry worksheetEntry = worksheets.get(iWorkSheet);
				URL listFeedUrl = worksheetEntry.getListFeedUrl();
				ListFeed wfeed = service.getFeed(listFeedUrl, ListFeed.class);
				for (ListEntry entry1 : wfeed.getEntries()) {
					System.out.println("Set to zero: "+entry1.getTitle().getPlainText());
					//if (entry1.getTitle().getPlainText().toLowerCase().contains("alpha")) {
						for (int k = 1; k < 8; k++) {
							entry1.getCustomElements().setValueLocal(daysOfTheWeek[k], "0");
						}
						entry1.update();
					//} else {
					//	entry1.delete();
					//}
				}
			//}

		}
	}
	public void RunTimePeriodReport(String theWeek,Integer iWorkSheet) throws Exception{
		URL metafeedUrl = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");
        SpreadsheetService service = new SpreadsheetService("EmailMeTime");
        service.setUserCredentials("troy.sadkowsky@gmail.com", "p34s0up");
        DocumentQuery query = new DocumentQuery(metafeedUrl);
        query.setTitleQuery("EmailMeTime");
		query.setTitleExact(true);
		query.setMaxResults(1);
				
        SpreadsheetFeed feed = service.getFeed(query, SpreadsheetFeed.class);
        List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        for (int i = 0; i < spreadsheets.size(); i++) {
          SpreadsheetEntry entry = spreadsheets.get(i);
          System.out.println("Found spreadsheet: "+entry.getTitle().getPlainText());
          List<WorksheetEntry> worksheets = entry.getWorksheets();
          if(worksheets.size()>1){
        	  WorksheetEntry worksheetEntry = worksheets.get(0);
        	  URL listFeedUrl = worksheetEntry.getListFeedUrl();
              ListFeed wfeed = service.getFeed(listFeedUrl, ListFeed.class); 
              boolean bFound = false;
              String dayOfTheWeek = "";
              String duration = "";
              String subject = "";
              for (ListEntry entry1 : wfeed.getEntries()) {
					for (String tag : entry1.getCustomElements().getTags()) {
						String thevalue = entry1.getCustomElements().getValue(tag);
						if (tag.equalsIgnoreCase("Week")) {
							if (thevalue.equalsIgnoreCase(theWeek)) {
								bFound = true;
								dayOfTheWeek = entry1.getCustomElements().getValue("DayOfWeek");
								duration = entry1.getCustomElements().getValue("Duration");
								subject = entry1.getCustomElements().getValue("Project");
								break;
							}
						}
					}
					if (bFound) {
						WorksheetEntry worksheetEntryPeriod = worksheets.get(iWorkSheet);
						URL listFeedUrlPeriod = worksheetEntryPeriod.getListFeedUrl();

						ListFeed wfeedPeriod = service.getFeed(listFeedUrlPeriod, ListFeed.class);
						boolean bFound1 = false;
						for (ListEntry entryPeriod : wfeedPeriod.getEntries()) {
							//System.out.println("Finding: "+entryPeriod.getTitle().getPlainText());
							for (String tag : entryPeriod.getCustomElements().getTags()) {
								String thevalue = entryPeriod.getCustomElements().getValue(tag);
								if (tag.equalsIgnoreCase("Project")) {
									if (thevalue.equalsIgnoreCase(subject)) {
										bFound1 = true;
										break;
									}
								}
							}
							if (bFound1) {
								String thevalue = entryPeriod.getCustomElements().getValue(dayOfTheWeek);
								if (thevalue != null) {
									float iValue = Float.valueOf(thevalue);
									float iNewValue = (float) 0.5;
									try {
										iNewValue = Float.valueOf(duration);
									} catch (Exception e) {
										// use default
									}
									float iTotalValue = iValue + iNewValue;
									entryPeriod.getCustomElements().setValueLocal(dayOfTheWeek,String.valueOf(iTotalValue));
								} else {
									entryPeriod.getCustomElements().setValueLocal(dayOfTheWeek,String.valueOf(duration));
								}
								System.out.println("Updating: "+entry1.getTitle().getPlainText());
								System.out.println("Updated: "+dayOfTheWeek + duration);
								entryPeriod.update();
								break;
							}
							bFound = false;
						}
					}
				}
			}
		}
	}
}
