import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainClass {

	private static Connection conn = null;
	
	//You can change this names according to your setup
	private static final String DATABASE_LOCATION = "databaseFile.db";
	private static final String EXCEL_FILE_LOCATION = "excelFile.xlsx";
	private static final String TABLE_NAME = "databaseTable";

	public static void main(String[] args) {
		try {
			conn = DatabaseConnector.getConnection(DATABASE_LOCATION);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try{
			openExcelAndWriteToDatabase();
		}catch(Exception ex) {
			ex.printStackTrace();
		}
	}

	private static void openExcelAndWriteToDatabase() throws Exception{
		//open file
		File excelFile = new File(EXCEL_FILE_LOCATION);
		FileInputStream fis = new FileInputStream(excelFile);

		//get sheet and prepare for getting cells
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);//because of 0, it just gets first sheet
		Iterator<Row> rowIterator = sheet.iterator();

		//IMPORTANT: First Row is special because it includes column names
		Row firstRow = rowIterator.next();
		Iterator<Cell> firstRowIterator = firstRow.cellIterator();

		ArrayList<String> columnNames = new ArrayList<String>();

		while(firstRowIterator.hasNext()) {
			columnNames.add(firstRowIterator.next().toString());
		}

		createDatabaseAndTable(columnNames);

		while(rowIterator.hasNext()) {
			Row row=rowIterator.next();
			Iterator <Cell> cellIterator= row.cellIterator();
			ArrayList<String> items = new ArrayList<String>();
			while(cellIterator.hasNext()) {
				items.add(cellIterator.next().toString());		
				//add all items in the row
			}
			writeToDatabase(items,columnNames);

		}
		fis.close();
		workbook.close();
		System.out.println("Excel Sheet is copied to SQLite Database successfully.");
	}

	private static void createDatabaseAndTable(ArrayList<String> columnNames) throws SQLException {
		//it is working for just one table. Run it several times for multiple tables (just considering first sheet)
		String sql = getCreateTableStatement(columnNames);
		Statement statement = conn.createStatement();
		statement.execute(sql);
	}


	private static void writeToDatabase(ArrayList<String> items, ArrayList<String> columnNames) throws SQLException {
		String query = getQueryStatement(columnNames);
		PreparedStatement pst = conn.prepareStatement(query);
		for(int i=1;i<=items.size();i++){
			pst.setString(i, items.get(i-1));
		}
		pst.execute();
		pst.close();

	}
	
	private static String getQueryStatement(ArrayList<String> columnNames) {
		//first and last lines of SQLite Command is special
		String query="insert into "+TABLE_NAME+" (";
		for(int i=0;i<columnNames.size()-1;i++) {
			query+=columnNames.get(i)+", ";
		}
		query+=columnNames.get(columnNames.size()-1)+") values (";

		for(int i=0;i<columnNames.size()-1;i++) {
			query+="?,";
		}
		query+="?)";
		
		return query;
	}

	private static String getCreateTableStatement(ArrayList<String> columnNames) {
		//first and last lines of SQLite Command is special
		String sql = "CREATE TABLE IF NOT EXISTS "+TABLE_NAME+" ("+columnNames.get(0)+" TEXT,";

		for(int i=1; i<columnNames.size()-1;i++) {
			sql+= " "+columnNames.get(i)+" TEXT,";
		}
		sql+= " "+columnNames.get(columnNames.size()-1)+" TEXT);";

		return sql;
	}
}
