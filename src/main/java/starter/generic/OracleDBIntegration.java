package starter.generic;

import org.apache.log4j.Logger;

import java.sql.*;
import java.util.HashMap;
import java.util.Map;

public class OracleDBIntegration {
    private static final Logger LOGGER = Logger.getLogger(OracleDBIntegration.class);
    private Connection con;
    private Statement stmt;
    private ResultSet rs;
    private static final String USERNAME ="";
    private static final String Passowrd ="";
    private static final String Url_JDBC ="jdbc:oracle:thin:@deheremclo2991.emea.adsint.biz:1521:aw11";
    public OracleDBIntegration() throws ClassNotFoundException, SQLException {
        Class.forName("oracle.jdbc.driver.OracleDriver");
        con = DriverManager.getConnection(Url_JDBC,USERNAME,Passowrd);
    }

    public void createStatment(String query,String queryMiddle,String queryEnd) throws SQLException {
        try{
            StringBuilder sqlQuery = new StringBuilder();
            sqlQuery.append(query).append(queryMiddle).append(queryEnd);
            stmt = con.createStatement();
            rs = stmt.executeQuery(sqlQuery.toString());
        } catch (SQLException e) {
            LOGGER.error("SQL Exception in create Statement",e);
            con.close();
        }
    }

    public String result(String column) throws SQLException {
        String p=null;
        int sum =0;
        while(rs.next()){
            p=rs.getString(column);
            if("Ordeno".equalsIgnoreCase(column)){
                sum+=rs.getInt(column);
            }
        }
        con.close();
        if("Ordeno".equalsIgnoreCase(column)){
            p=""+sum;
        }
        return p;
    }

    public  void updateStatement(String query) throws SQLException {
        try{
            stmt=con.createStatement();
            stmt.executeUpdate(query);
            con.commit();
            con.close();
        } catch (SQLException e) {
            LOGGER.error("SQL Exception in update Statement",e);
            con.close();
        }
    }

    public Map<String,String> resultWithNullCheck(String column) throws SQLException {
        Map<String,String> col =new HashMap<>();
        ResultSetMetaData columns = rs.getMetaData();

        while(rs.next()){
            int colcount =columns.getColumnCount();
            for (int i=1;i<=colcount;i++){
                String value =rs.getString(i);
                if(rs.getString(i)!=null){
                    col.put(columns.getCatalogName(i).toLowerCase(),value);
                }
                else{
                    col.put(columns.getCatalogName(i).toLowerCase(),"");
                }
            }

        }
        con.close();

        return col;
    }
}
