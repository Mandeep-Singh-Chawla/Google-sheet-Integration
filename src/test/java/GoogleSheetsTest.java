

import java.io.IOException;
import java.util.List;

import com.google.api.services.sheets.v4.model.ValueRange;

import CommonTexts.Constant;
import utility.ConnectToGSheet;


public class GoogleSheetsTest{
	
	List<List<Object>> sheetRows;
	
	public GoogleSheetsTest() throws IOException {
		// TODO Auto-generated constructor stub
		ConnectToGSheet.initGoogleWorkBook();
	}

	public void readData() throws Exception {
		ValueRange response = ConnectToGSheet.initGoogleSheet(Constant.GSheetTest,
				Constant.CalculationSheetRange, Constant.MajorDimension_Row);

		sheetRows = ConnectToGSheet.getRows(response);
		System.out.println(sheetRows.toString());
	}
	
	public void updateData() throws IOException {
		for(List<Object> obj:sheetRows) {
			
			int add = Integer.parseInt(obj.get(0).toString()) + Integer.parseInt(obj.get(1).toString());
			int multi = Integer.parseInt(obj.get(0).toString()) * Integer.parseInt(obj.get(1).toString());
			int sub = Integer.parseInt(obj.get(0).toString()) - Integer.parseInt(obj.get(1).toString());
			
			obj.add(add);
			obj.add(multi);
			obj.add(sub);
			
		}
		
		ConnectToGSheet.updateDataRowsToSheet(Constant.GSheetTest, Constant.CalculationSheetRange, sheetRows, "USER_ENTERED");
		
	}
	
	public static void main(String[] args) throws Exception {
		GoogleSheetsTest gSheet= new GoogleSheetsTest();
		gSheet.readData();
		gSheet.updateData();
	}
}
