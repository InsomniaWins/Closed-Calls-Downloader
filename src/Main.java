import static java.util.concurrent.TimeUnit.*;

import org.json.*;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;

public class Main {
	
	private static final String ORIGINAL_SAVE_NAME = "Output.xls";
	
	public static void main(String[] args) throws MalformedURLException {
		
		final ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);
		
		URL url = new URL("https://p2c.beaumonttexas.gov/p2c/cad/cadHandler.ashx?op=s");
		
		final Runnable checkLoop = new Runnable() {
			public void run() {
				
				System.out.println("Checking for new content . . .");
				
				String postData = "t=css&_search=false&nd=" + Long.toString(System.currentTimeMillis()) + "&rows=10000&page=1&sidx=starttime&sord=desc";
				
				URLConnection connection;
				try {
					connection = url.openConnection();
					connection.setDoOutput(true);
					connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
					connection.setRequestProperty("Content-Length", Integer.toString(postData.length()));
					
					DataOutputStream dos = new DataOutputStream(connection.getOutputStream());
					dos.writeBytes(postData);
					
					BufferedReader buffReader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
					String line;
					String jsonString = "";
					while ((line = buffReader.readLine()) != null) {
						if (!jsonString.isEmpty()) jsonString = jsonString + "\n";
						jsonString = jsonString + line;
					}

					JSONObject json = new JSONObject(jsonString);
					String rowsString = json.get("rows").toString();
					
					WritableWorkbook workbook = null;
					
					if (new File(ORIGINAL_SAVE_NAME).exists()) {
						try {
							Workbook oldWorkbook = Workbook.getWorkbook(new File(ORIGINAL_SAVE_NAME));
							workbook = Workbook.createWorkbook(new File(ORIGINAL_SAVE_NAME), oldWorkbook);
						} catch (FileNotFoundException e) {
							System.out.println("Could not open file! (File Is In Use By Another Program)");
							System.exit(0);
						}
						
					} else {
						workbook = Workbook.createWorkbook(new File(ORIGINAL_SAVE_NAME));
						workbook.createSheet("Test Sheet", 0);
					}
					
					WritableSheet sheet = workbook.getSheet(0);
					
					// create labels
					Label agencyLabel = new Label(0,0,"Agency");
					Label serviceLabel = new Label(1,0,"Service");
					Label caseNumberLabel = new Label(2,0,"Case Number");
					Label startTimeLabel = new Label(3,0,"Start Time");
					Label endTimeLabel = new Label(4,0,"End Time");
					Label natureLabel = new Label(5,0,"Nature");
					Label addressLabel = new Label(6,0,"Address");
					
					sheet.addCell(agencyLabel);
					sheet.addCell(serviceLabel);
					sheet.addCell(caseNumberLabel);
					sheet.addCell(startTimeLabel);
					sheet.addCell(endTimeLabel);
					sheet.addCell(natureLabel);
					sheet.addCell(addressLabel);
					
					
					for (int i = 0; i < rowsString.length(); i++) {
						if ( rowsString.charAt(i) == '{') {
							
							int offset = 1;
							while (rowsString.charAt(i + offset) != '}') {
								offset += 1;
							}
							
							String rowData = rowsString.substring(i, i+offset);
							
							String agency = getStringValueFromString("agency",rowData);
							String service = getStringValueFromString("service",rowData);
							String caseNumber = getStringValueFromString("id",rowData);
							String startTime = getStringValueFromString("starttime",rowData);
							String endTime = getStringValueFromString("closetime",rowData);
							String nature = getStringValueFromString("nature",rowData);
							String address = getStringValueFromString("address",rowData);

							sheet.setColumnView(0, 10);
							sheet.setColumnView(1, 10);
							sheet.setColumnView(2, 20);
							sheet.setColumnView(3, 26);
							sheet.setColumnView(4, 26);
							sheet.setColumnView(5, 34);
							sheet.setColumnView(6, 60);
							
							int rowIndex = sheetHasCase(sheet, caseNumber);
							
							if (rowIndex != -1) {
								agencyLabel = new Label(0,rowIndex, agency);
								serviceLabel = new Label(1,rowIndex, service);
								caseNumberLabel = new Label(2,rowIndex, caseNumber);
								startTimeLabel = new Label(3,rowIndex, startTime);
								endTimeLabel = new Label(4,rowIndex, endTime);
								natureLabel = new Label(5,rowIndex, nature);
								addressLabel = new Label(6,rowIndex,address);
							} else {
								rowIndex = sheet.getRows();
								agencyLabel = new Label(0,rowIndex, agency);
								serviceLabel = new Label(1,rowIndex, service);
								caseNumberLabel = new Label(2,rowIndex, caseNumber);
								startTimeLabel = new Label(3,rowIndex, startTime);
								endTimeLabel = new Label(4,rowIndex, endTime);
								natureLabel = new Label(5,rowIndex, nature);
								addressLabel = new Label(6,rowIndex,address);
								String callString = "{ agency:"+agency+" service:"+service+" case number:"+caseNumber+" start time:"+startTime+" end time:"+endTime+" nature: "+nature+" address"+address+" }";
								
								System.out.println("New Call Found! "+callString);
							}
							
							sheet.addCell(agencyLabel);
							sheet.addCell(serviceLabel);
							sheet.addCell(caseNumberLabel);
							sheet.addCell(startTimeLabel);
							sheet.addCell(endTimeLabel);
							sheet.addCell(natureLabel);
							sheet.addCell(addressLabel);
						}
						
					}
					
					workbook.write();
					workbook.close();
					
					
				} catch (IOException | WriteException | BiffException e) {
					e.printStackTrace();
				}
				
				System.out.println("Next check will be in 10 hours.");
				
				
			}
		};
		
		scheduler.scheduleAtFixedRate(checkLoop, 0, 10, HOURS);
		
	}
	
	private static int sheetHasCase(WritableSheet sheet, String caseNumber) {
		for (int i = 0; i < sheet.getRows(); i++) {
			Label cell = (Label) sheet.getCell(2,i);
			if (cell.getString().contains(caseNumber)) return i;
		}
		
		return -1;
	}
	
	private static int[] findSubstringInString(String substring, String string) {
		int[] returnInt = {-1,-1};
		
		for (int i = 0; i < string.length(); i++) {
			
			if (string.charAt(i) == substring.charAt(0) && i + substring.length() <= string.length()) {
				
				String substringCheck = string.substring(i, i+substring.length());
				
				if (substringCheck.equals(substring)) {
					returnInt[0] = i;
					returnInt[1] = i+substring.length();
					break;
					
				}
				
			}
			
		}
		
		return returnInt;
	}
	
	private static int[] findValuePoints(String key, String table) {
		
		int[] returnArray = {0,0};
		
		returnArray = findSubstringInString("\""+key+"\"", table);
		returnArray[1] += 2;
		returnArray[0] = returnArray[1];
		
		while (table.charAt(returnArray[1]) != '"') {
			returnArray[1] += 1;
		}
		
		return returnArray;
		
	}
	
	private static String getStringValueFromString(String key, String table) {
		
		int valuePoints[] = findValuePoints(key, table);
		
		return table.substring(valuePoints[0], valuePoints[1]);
	}
	
}	
