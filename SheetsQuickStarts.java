# csc131
attendanceTracker
package SheetPackageTest;

// This is SheetsQuickstart. This program is used as a demo to demostrate an update to a fix cell location on a Google sheet
// Todos: Please be sure to update the location of your client_secret.json file & the Googlesheet id before running your program.
// Author: Doan Nguyen
// Date: 6/1/17
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetsQuickstart {
	/** Application name. */
	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";

	/** Directory to store user credentials for this application. */
	private static final java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"),
			".credentials//sheets.googleapis.com-java-quickstart.json");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	/**
	 * Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials at
	 * ~/.credentials/sheets.googleapis.com-java-quickstart.json
	 */
	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);
	private static final List<String> READER = Arrays.asList(SheetsScopes.SPREADSHEETS_READONLY);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	public static Credential authorize() throws IOException {
		// Load client secrets.
		// Todo: Change this text to the location where your client_secret.json resided
		InputStream in = new FileInputStream(
				"C:\\Users\\chukwuemelie\\Downloads\\working directory\\client_secret.json");
		// SheetsQuickstart.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES).setDataStoreFactory(DATA_STORE_FACTORY).setAccessType("offline").build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	public static Credential authorizeRead() throws IOException {
		// Load client secrets.
		InputStream in = new FileInputStream(
				"C:\\\\Users\\\\chukwuemelie\\\\Downloads\\\\working directory\\\\client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, READER).setDataStoreFactory(DATA_STORE_FACTORY).setAccessType("offline").build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * 
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	public static Sheets getSheetsService() throws IOException {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).setApplicationName(APPLICATION_NAME)
				.build();
	}

	public static boolean validKey(int key) {
		int keyID = 1234;
		if (key == keyID)
			return true;
		else
			return false;
	}

	// checks instructor key
	public static boolean isInstructor(int instructorID) {
		int profID = 123456789;

		if (instructorID == profID) {
			return true;
		}
		return false;
	}

	// checks student key
	public static boolean isStudent(int sid) throws IOException {
		Sheets service = getSheetsService();
		String spreadsheetId = "1NZzlnRo5jKemmfdPf3H6na15Fwe4-fD5ZI7Bvs3W_GI";
		String checkID = Integer.toString(sid);
		String range = "Sheet1!A2:G";

		ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();

		List<List<Object>> cellID = response.getValues();

		if (cellID != null && cellID.size() != 0) {
			for (List row : cellID) {
				// Check if input ID is in spreadsheets
				if (checkID.equals(row.get(1))) {
					return true;
				}
			}
		}
		return false;
	}

	public static void recordKey(String key /* , int dateCol, int dayCount */) throws IOException {
		int dateCol = 3;
		int dayCount = 0;
		Sheets service = getSheetsService();
		String spreadsheetId = "1NZzlnRo5jKemmfdPf3H6na15Fwe4-fD5ZI7Bvs3W_GI";
		List<Request> requests = new ArrayList<>();
		List<CellData> inputKey = new ArrayList<>();

		inputKey.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue((key))));

		// Prepare request with proper row and column and its value
		requests.add(new Request().setUpdateCells(new UpdateCellsRequest()
				.setStart(new GridCoordinate().setSheetId(0).setRowIndex(16) // set the row to row 0
						.setColumnIndex(dateCol + (dayCount - 1)))
				.setRows(Arrays.asList(new RowData().setValues(inputKey)))
				.setFields("userEnteredValue,userEnteredFormat.backgroundColor")));

		BatchUpdateSpreadsheetRequest batchUpdateRequestKey = new BatchUpdateSpreadsheetRequest().setRequests(requests);
		service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequestKey).execute();
	}

	
	public static void updateSheet(int sid, int key) throws IOException {
		// Build a new authorized API client service.
		Sheets service = getSheetsService();

		String spreadsheetId = "13_b-hy1dCToPbDntmTFb2Lu5kfcnOrh9KZ9ymVLCygY";

		// Create requests object
		List<Request> requests = new ArrayList<>();

		// Create values object
		List<CellData> values = new ArrayList<>();

		// Getting current date and creates new column if it is a new date
		String dateRange = "Sheet1!C1:Z1";
		ValueRange dateResponse = service.spreadsheets().values().get(spreadsheetId, dateRange).execute();
		List<List<Object>> dateCellID = dateResponse.getValues();

		int dateCol = 2;
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM/dd/yyyy");
		String date = dtf.format(LocalDate.now());

		// Counts the number of dates used
		String dayRange = "Sheet1!B2:B16";
		ValueRange dayResponse = service.spreadsheets().values().get(spreadsheetId, dayRange).execute();
		List<List<Object>> dayCellID = dayResponse.getValues();

		int dayCount = 0;
		String dayCountString = "";
		for (List dayRow : dayCellID) {
			dayCountString = (String) dayRow.get(0);
			dayCount = Integer.parseInt(dayCountString);
		}
		if (dayCount == 0) {
			dayCount++;
			dayCountString = Integer.toString(dayCount);
			List<CellData> days = new ArrayList<>();
			days.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue((dayCountString))));

			requests.add(new Request().setUpdateCells(new UpdateCellsRequest()
					.setStart(new GridCoordinate().setSheetId(0).setRowIndex(15) // set the row to row 0
							.setColumnIndex(1)) // set the new column 6 to value date at row 0
					.setRows(Arrays.asList(new RowData().setValues(days)))
					.setFields("userEnteredValue,userEnteredFormat.backgroundColor")));
		} else {
			for (List row : dateCellID) {
				if (!(date.equals(row.get(dayCount - 1))))
					dayCount++;
			}
			dayCountString = Integer.toString(dayCount);
			List<CellData> days = new ArrayList<>();
			days.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue((dayCountString))));

			requests.add(new Request().setUpdateCells(new UpdateCellsRequest()
					.setStart(new GridCoordinate().setSheetId(0).setRowIndex(15) // set the row to row 0
							.setColumnIndex(1)) // set the new column 6 to value date at row 0
					.setRows(Arrays.asList(new RowData().setValues(days)))
					.setFields("userEnteredValue,userEnteredFormat.backgroundColor")));
		}

		// Sets columns for date

		values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue((date))));

		requests.add(new Request().setUpdateCells(new UpdateCellsRequest()
				.setStart(new GridCoordinate().setSheetId(0).setRowIndex(0) // set the row to row 0
						.setColumnIndex(dateCol + (dayCount - 1))) // set the new column 6 to value date at row 0
				.setRows(Arrays.asList(new RowData().setValues(values)))
				.setFields("userEnteredValue,userEnteredFormat.backgroundColor")));

		// Put passcode in excel
		String pkey = Integer.toString(key);
		// recordKey(pkey, dateCol, dayCount);

		BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest().setRequests(requests);
		service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();

		// checks to see if student id is in the list
		String checkID = Integer.toString(sid);
		String range = "Sheet1!B2:G";
		if (validKey(key)) {
			ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
			List<List<Object>> cellID = response.getValues();
			if (cellID != null && cellID.size() != 0) {
				int rowCount = 1;
				for (List row : cellID) {
					// Check if input ID is in spreadsheets
					if (checkID.equals(row.get(0))) {
						List<CellData> valuesNew = new ArrayList<>();
						// Add string "Present value
						valuesNew.add(
								new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(("Present"))));

						// Prepare request with proper row and column and its value
						requests.add(new Request().setUpdateCells(new UpdateCellsRequest()
								.setStart(new GridCoordinate().setSheetId(0).setRowIndex(rowCount++) // set the row to
																										// row 1
										.setColumnIndex(dateCol + (dayCount - 1))) // Marks "Present" on the same column
																					// as the current date
								.setRows(Arrays.asList(new RowData().setValues(valuesNew)))
								.setFields("userEnteredValue,userEnteredFormat.backgroundColor")));
					}
					rowCount++;
				}
			}
			BatchUpdateSpreadsheetRequest batchUpdateRequestNew = new BatchUpdateSpreadsheetRequest()
					.setRequests(requests);
			service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequestNew).execute();
		}

	}

}
