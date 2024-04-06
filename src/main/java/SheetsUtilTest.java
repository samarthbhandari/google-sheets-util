import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class SheetsUtilTest {
    private static final String APPLICATION_NAME = "Test";
    private static final String ROLE = "writer";
    private static final String STARTINGRANGE = "A1";
    private static final String RAW = "RAW";
    private static final String SHEETNAME = "Test sheet ";

    private static final String SERVICE_ACCOUNT = """
            """;

    public static void main(String[] args) {
        List<String> users = new ArrayList<>();
        users.add("test@gmail.com");
        users.add("test2@gmail.com");
        List<List<Object>> values = new ArrayList<>();
        List<Object> row = new ArrayList<>();
        row.add("test1");
        row.add("test2");
        row.add("test3");
        values.add(row);
        row = new ArrayList<>();
        row.add("test4");
        row.add("test5");
        row.add("test6");
        values.add(row);
        List<String> sheets = new ArrayList<>();
        sheets.add("Sheet 1");
        sheets.add("Sheet 2");
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(new Date());
        String finalSheetName = SHEETNAME + calendar.get(Calendar.YEAR) + "-" + (calendar.get(Calendar.MONTH) + 1) + "-" + calendar.get(Calendar.DAY_OF_MONTH);
        GoogleSheetsUtil util = new GoogleSheetsUtil(APPLICATION_NAME, SERVICE_ACCOUNT);
        String createdWorkBookId = util.findWorkBookWithName("Test Sheet");
        System.out.println(createdWorkBookId);
    }
}
