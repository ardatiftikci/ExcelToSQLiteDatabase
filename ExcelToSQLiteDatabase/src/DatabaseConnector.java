import java.sql.Connection;
import java.sql.DriverManager;

public class DatabaseConnector {
		
	public static Connection getConnection(String location) throws Exception{
		try {
			Class.forName("org.sqlite.JDBC");
			Connection conn = DriverManager.getConnection("jdbc:sqlite:"+location);
			return conn;
		}catch(Exception e) {
			e.printStackTrace();
			return null;
		}
	}
}
