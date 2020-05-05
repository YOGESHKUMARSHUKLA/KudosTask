package starter.generic;

public class CsvMapper extends InputTemplate {
    private static final String CSV_LOCATION = "/template/csv/";
    public CsvMapper(String path){super(CSV_LOCATION+path);}
}
