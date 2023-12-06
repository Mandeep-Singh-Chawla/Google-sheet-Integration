package utility;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Append;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.BatchUpdate;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Clear;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Values.Update;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.ClearValuesResponse;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import CommonTexts.ConstantLiterals;

public final class ConnectToGSheet {

	private static final java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"), ConstantLiterals.GSHEETCREDENTIALSPATH);
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);
	private static FileDataStoreFactory DATA_STORE_FACTORY;
	private static HttpTransport HTTP_TRANSPORT;
	private static ValueRange response;
	private static Sheets service;
	private final static  Logger log = LogManager.getLogger(ConnectToGSheet.class.getName());

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} 
		catch (Throwable t) {
			System.exit(1);
		}
	}

	/**
	 * Creates an authorized Credential object.
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	private static Credential authorize() throws IOException {
		InputStream in = ConnectToGSheet.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(DATA_STORE_FACTORY)
				.setAccessType("offline")
				.build();

		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		return credential;
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	private static Sheets getSheetsService() throws IOException {
		Credential credential;
		credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
				.setApplicationName(ConstantLiterals.APPLICATION_NAME)
				.build();
	}

	/**
	 * Build and return values of the cells defined in the range.
	 * @return a list of list of objects
	 * @throws IOException
	 */
	public static void initGoogleWorkBook() throws IOException {
		service = getSheetsService();
		if(service != null) {
			log.info("Google workbook initialized");
		}
	}

	public static ValueRange initGoogleSheet(String spreadsheetId, String range, String majorDimensions) throws IOException {
		return response = service.spreadsheets()
				.values()
				.get(spreadsheetId, range)
				.setMajorDimension(majorDimensions)
				.execute();
	}

	private static List<List<Object>> setValue(Object value) {
		return Arrays.asList(Arrays.asList(value));
	}

	public static List<List<Object>> getRows(ValueRange res) throws IOException {
		return res.getValues();
	}

	public static int getGoogleSheetSize() {
		int size = response.getValues().size();
		return size;
	}

	//TODO Delete this method, new methods are updateCellValues() and appendCellValues()
	public static void setCellValue(String spreadsheetId, String range, String value) throws IOException {

		ValueRange requestBody = new ValueRange();
		requestBody.setValues(setValue(value));
		Update request = service.spreadsheets().values().update(spreadsheetId, range, requestBody);
		request.setIncludeValuesInResponse(Boolean.TRUE);
		request.setValueInputOption("RAW");

		UpdateValuesResponse response = request.execute();
		log.info("setCellValue: "+response);
	}

	public static void clearCellValue(String spreadSheetId, String range, ClearValuesRequest requestBody) throws IOException {
		Clear request = service.spreadsheets().values().clear(spreadSheetId, range, requestBody);
		ClearValuesResponse response = request.execute();
		log.info("clearCellValue: "+response);
	}
	
	private static void updateCellValues(String spreadsheetId, String range, ValueRange requestBody) throws IOException {
		Update request = service.spreadsheets().values().update(spreadsheetId, range, requestBody);
		request.setIncludeValuesInResponse(Boolean.TRUE);
		request.setValueInputOption("RAW");

		UpdateValuesResponse response = request.execute();
		log.info("updateCellValues: "+response);
	}

	private static void appendCellValues(String spreadsheetId, String range, ValueRange requestBody) throws IOException {
		Append request = service.spreadsheets().values().append(spreadsheetId, range, requestBody);
		request.setValueInputOption("RAW");

		AppendValuesResponse response = request.execute();
		log.info("appendCellValues: "+response);
	}

	public static void batchUpdateDataValue(String spreadsheetId, List<ValueRange> data) throws IOException {
		BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
		requestBody.setValueInputOption("RAW");
		requestBody.setData(data);
		BatchUpdate request = service.spreadsheets().values().batchUpdate(spreadsheetId, requestBody);

		BatchUpdateValuesResponse response = request.execute();
		log.info("Batch updated: "+response);
	}

	/**
	 * Build and return values of the cells defined in the range.
	 * @return a list of list of objects
	 * @throws IOException
	 */
	public static List<List<Object>> getCellValues(String spreadsheetId, String range, String majorDimensions) throws IOException {
		Sheets service = getSheetsService();
		try {
			if(service != null) {
				ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).setMajorDimension(majorDimensions).execute();
				List<List<Object>> listListObject = response.getValues();

				if (listListObject == null || listListObject.size() == 0) {
					return null;
				} 
				else {

					log.info(response.toPrettyString());
					return listListObject;
				}
			}
		}
		catch (IOException e) {
			log.info("getCellValues: "+e.getMessage());
			e.printStackTrace();
		}
		return null;
	}

	public static Object getCellValue(int row, int column) throws NullPointerException{
		try {
			return response.getValues().get(row-1).get(column-1);
		}
		catch (IndexOutOfBoundsException e) {
			return "";
		}
	}
	
	public static void updateDataToSheet(String spreadSheetId, String range, List<Object> data, String majorDimension) {
		try {
			ValueRange requestBody = new ValueRange();
			requestBody.setValues(Arrays.asList(data));
			requestBody.setMajorDimension(majorDimension);
			requestBody.setRange(range);
			updateCellValues(spreadSheetId, range, requestBody);
		} 
		catch (IOException e) {
			log.info("updateDataToSheet: "+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void updateDataRowsToSheet(String spreadSheetId, String range, List<List<Object>> data, String majorDimension) {
		try {
			
			ValueRange body = new ValueRange()
			        .setValues(data);
			
			UpdateValuesResponse result =
			        service.spreadsheets().values().update(spreadSheetId,range, body)
			                .setValueInputOption(majorDimension)
			                .execute();
			System.out.printf("%d cells updated.", result.getUpdatedCells());
		} 
		catch (IOException e) {
			log.info("updateDataToSheet: "+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void appendDataToSheet(String spreadSheetId, String range, List<Object> data, String majorDimension) {
		try {
			ValueRange requestBody = new ValueRange();
			requestBody.setValues(Arrays.asList(data));
			requestBody.setMajorDimension(majorDimension);
			requestBody.setRange(range);
			appendCellValues(spreadSheetId, range, requestBody);
		} 
		catch (IOException e) {
			log.info("appendDataToSheet: "+e.getMessage());
			e.printStackTrace();
		}
	}
	
	public static void clearDataFromSheet(String spreadSheetId, String range) throws IOException {
		ClearValuesRequest requestBody = new ClearValuesRequest();
		Clear request = service.spreadsheets().values().clear(spreadSheetId, range, requestBody);
		ClearValuesResponse response = request.execute();
		log.info("clearDataFromSheet: "+response);
	}

	//TODO
	public static void updateData(String atRow, String spreadSheetId, String range, List<Object> data, String majorDimension) {
		try {
			ValueRange requestBody = new ValueRange();
			requestBody.setValues(Arrays.asList(data));
			requestBody.setMajorDimension(majorDimension);
			updateCellValues(spreadSheetId, range, requestBody);
		} 
		catch (IOException e) {
			log.info("updateData: "+e.getMessage());
			e.printStackTrace();
		}
	}

	public static  Map<String, Character> toFromOfColumns(String range) {
		Map<String, Character> info	= new HashMap<String, Character>();
		info.put("to",  '\u0000');
		info.put("from", '\u0000');
		range						= range.substring(range.indexOf("!")+1);
		if(range.contains(":")) {
			String[] characters		= range.split(":");
			info.put("to", characters[0].replaceAll("\\d*", "").toCharArray()[0]);
			info.put("from", characters[1].replaceAll("\\d*", "").toCharArray()[0]);
		}
		else {
			info.put("to", range.replaceAll("\\d*", "").toCharArray()[0]);
			info.put("from", '\u0000');
		}
		return info;
	}
	
	
}

