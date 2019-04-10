package excelparser;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;


public class App 
{	
	
	private static String[] slices = {"00:00", "0:15","0:30","0:45","1:00","1:15","1:30","1:45","2:00","2:15","2:30","2:45","3:00","3:15","3:30","3:45","4:00","4:15","4:30","4:45",
			  "5:00","5:15","5:30","5:45","6:00","6:15","6:30","6:45","7:00","7:15","7:30","7:45","8:00","8:15","8:30","8:45","9:00","9:15","9:30",
			  "9:45","10:00","10:15","10:30","10:45","11:00","11:15","11:30","11:45","12:00","12:15","12:30","12:45","13:00","13:15","13:30","13:45",
			  "14:00","14:15","14:30","14:45","15:00","15:15","15:30","15:45","16:00","16:15","16:30","16:45","17:00","17:15","17:30","17:45","18:00",
			  "18:15","18:30","18:45","19:00","19:15","19:30","19:45","20:00","20:15","20:30","20:45","21:00","21:15","21:30","21:45","22:00","22:15",
			  "22:30","22:45","23:00","23:15","23:30","23:45"};
	
	private static String[] days = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
	
    public static void main( String[] args ) throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException
    {
    	
    	
    	String curDir 					= System.getProperty("user.dir");
    	File folder 					= new File(curDir);
    	List<File> files				= listFilesInFolder(folder);
    	BufferedReader br				= null;
    	String line						= "";
    	List<String> l					= null;
    	SimpleDateFormat sdf			= new SimpleDateFormat("mm/dd/yyyy");
    	ArrayList<Interval> intervals 	= new ArrayList<Interval>();


    	// step through files and process
		for (File file : files) {
			System.out.println("Processing: "+file.getName());
			
			
			// Creating a Workbook from an Excel file (.xls or .xlsx)
	        Workbook workbook = WorkbookFactory.create(file);

	        // Retrieving the number of sheets in the Workbook
	        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
			
	        // Getting the Sheet at index zero
	        Sheet sheet = workbook.getSheetAt(0);

	        // Create a DataFormatter to format and get each cell's value as String
	        DataFormatter dataFormatter = new DataFormatter();
			
	        
	        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
	        boolean firstRow = true;
	        for (Row row : sheet) {
	        	
	        	boolean firstCell = true;
	        	Date day = null;
	        	String dow = null;
	        	int n = 0;
	        	
	            for(Cell cell: row) {
	            	
	            	if (firstRow) {
		        		continue;
		        	}
	            	
	                String cellValue = dataFormatter.formatCellValue(cell);
	                
	                if (firstCell) {
	            		day = sdf.parse(cellValue);
	            		dow = days[day.getDay()];
	            		firstCell = false;
	            		continue;
	            	}
	                
	                String d = sdf.format(day) +" "+ slices[n];
	                
	                Interval i = new Interval(d, dow, Double.parseDouble(cellValue));
	                
	                intervals.add(i);
	                
	                //System.out.print(cellValue + "\t");
	                n++;
	            }
	            firstRow = false;
	        }
			
	        System.out.println(intervals.size());
	        
	        // Closing the workbook
	        workbook.close();
	        
	        
	        try {
				PrintWriter pw 		= new PrintWriter(new File("formatted_"+file.getName()+".csv"));
				StringBuilder sb	= new StringBuilder();
				
				sb.append("Timestamp,Day,Value\n");
				
				for (Interval i : intervals) {
					sb.append(i.getTimestamp()+","+i.getDOW()+","+i.getValue()+"\n");
				}
				
				pw.write(sb.toString());
				pw.close();
			} catch(Exception e) {
				e.printStackTrace();
			}
		}
    }
    
    
    /**
	 * List all Excel (xls, xlsx) files in a folder
	 * @param File folder
	 * @return List<File>
	 */
	public static List<File> listFilesInFolder(final File folder) {
		List<File> files = new ArrayList<File>();
		for (final File fileEntry : folder.listFiles()) {
			if (fileEntry.isDirectory()) {
				// listFilesInFolder(fileEntry); // recursively list files in sub-folders
			} else if (fileEntry.getName().endsWith(".xls") || fileEntry.getName().endsWith(".xlsx")) {
				files.add(fileEntry);
			}
		}
		return files;
	}
}


/**
 * Object to store interval data
 */
class Interval {
	String timestamp 	= null;
	String dow			= null;
	Double value		= null;
	
	Interval(String t, String d, Double v) {
		this.timestamp = t;
		this.dow = d;
		this.value = v;
	}
	
	public String getTimestamp() 	{  return this.timestamp;  }
	public String getDOW() 			{  return this.dow;  }
	public Double getValue() 		{  return this.value;  }
}
