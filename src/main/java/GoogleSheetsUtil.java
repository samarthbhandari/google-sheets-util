import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class GoogleSheetsUtil {

    private static final Logger logger = LogManager.getLogger(GoogleSheetsUtil.class);

    private static final String SERVICE_ACCOUNT_FILE_PATH = "service-account.json";
    private final String applicationName;
    private final String serviceAccountJsonString;

    GoogleSheetsUtil(String applicationName) {
        this.applicationName = applicationName;
        this.serviceAccountJsonString = null;
    }

    GoogleSheetsUtil(String applicationName, String serviceAccountJsonString) {
        this.applicationName = applicationName;
        this.serviceAccountJsonString = serviceAccountJsonString;
    }

    public String getApplicationName() {
        return applicationName;
    }

    public String getServiceAccountJsonString() {
        return serviceAccountJsonString;
    }

    @SuppressWarnings("deprecation")
    private GoogleCredential getCredentials(Collection<String> scopes) {
        logger.debug("Getting credentials for scopes " + scopes);
        try {
            return GoogleCredential.fromStream(
                            Objects.requireNonNull(getInputStreamForServiceAccount()))
                    .createScoped(scopes);
        }
        catch (Exception e) {
            logger.error("Exception occurred in getCredentials method " + e.getMessage(), e);
            return null;
        }
    }
    private InputStream getInputStreamForServiceAccount() {
        logger.debug("Inside method getInputStreamForServiceAccount");
        InputStream inputStream = null;
        try {
            if (getServiceAccountJsonString() != null && getServiceAccountJsonString().isEmpty())
                inputStream = new ByteArrayInputStream(getServiceAccountJsonString().getBytes(StandardCharsets.UTF_8));
            else
                inputStream = GoogleSheetsUtil.class.getResourceAsStream(SERVICE_ACCOUNT_FILE_PATH);
        }
        catch (Exception e) {
            logger.error("Exception occurred in getInputStreamForServiceAccount method " + e.getMessage(), e);
        }
        return inputStream;
    }

    private Sheets initializeSheetService() {
        logger.debug("Inside method createSheetService");
        try {
            return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
                    GsonFactory.getDefaultInstance(),
                    getCredentials(Collections.singleton(SheetsScopes.SPREADSHEETS)))
                    .setApplicationName(getApplicationName())
                    .build();
        }
        catch (Exception e) {
            logger.error("Exception occurred in initializing sheets service " + e.getMessage(), e);
            return null;
        }
    }

    private Drive initializeDriveService() {
        logger.debug("Initializing drive service");
        try {
            return new Drive.Builder(GoogleNetHttpTransport.newTrustedTransport(),
                    GsonFactory.getDefaultInstance(),
                    getCredentials(Collections.singleton(DriveScopes.DRIVE_FILE)))
                    .setApplicationName(getApplicationName())
                    .build();
        }
        catch (Exception e) {
            logger.error("Exception occurred in initializing drive service " + e.getMessage(), e);
            return null;
        }
    }

    public void shareWorkBook(String workBookId, String user, String role) {
        List<String> users = new ArrayList<>();
        users.add(user);
        shareWorkBook(workBookId, users, role);
    }

    public void shareWorkBook(String workBookId, List<String> users, String role) {
        logger.debug("Sharing workbook with users : " + users);
        try {
            Drive driveService = initializeDriveService();
            if (driveService != null) {
                users.forEach(user -> {
                    Permission userPermission = new Permission()
                            .setType("user")
                            .setRole(role)
                            .setEmailAddress(user);
                    try {
                        driveService
                                .permissions()
                                .create(workBookId, userPermission)
                                .execute();
                    } catch (Exception e) {
                        logger.error("Exception occurred in sharing workbook " + workBookId + " to user " + user + " " + e.getMessage(), e);
                    }
                });
            }
            else
                logger.error("drive service is null, hence sharing failed");
        }
        catch (Exception e) {
            logger.error("Exception occurred in sharing workbook " + e.getMessage(), e);
        }
    }

    public String createWorkBookWithDefaultSheet(String spreadSheetName) {
        logger.debug("Inside method createWorkBookWithDefaultSheet");
        String workBookId = "";
        try {
            Sheets sheetService = initializeSheetService();
            if (sheetService != null) {
                Spreadsheet spreadsheet = sheetService.spreadsheets()
                        .create(
                                new Spreadsheet()
                                .setProperties(new SpreadsheetProperties()
                                        .setTitle(spreadSheetName))
                        )
                        .setFields("spreadsheetId")
                        .execute();
                workBookId = spreadsheet.getSpreadsheetId();
            }
            else
                logger.debug("sheetService is null. Failed to create spreadsheet");
        }
        catch (Exception e) {
            logger.error("Exception occurred in createWorkBookWithDefaultSheet method " + e.getMessage(), e);
        }
        logger.debug("Got workbook with id " + workBookId);
        return workBookId;
    }

    public String createWorkBookWithCustomSheets(String spreadSheetName, String sheetName) {
        List<String> sheetNames = new ArrayList<>();
        sheetNames.add(sheetName);
        return createWorkBookWithCustomSheets(spreadSheetName, sheetNames);
    }

    public String createWorkBookWithCustomSheets(String spreadSheetName, List<String> sheetNames) {
        logger.debug("Inside method createWorkBookWithDefaultSheet");
        String workBookId = "";
        try {
            Sheets sheetService = initializeSheetService();
            if (sheetService != null) {
                Spreadsheet workbook = new Spreadsheet()
                        .setProperties(new SpreadsheetProperties().setTitle(spreadSheetName));
                List<Request> sheets = new ArrayList<>();
                sheetNames.forEach(sheet -> {
                    logger.debug("Adding sheet " + sheet);
                    sheets.add(new Request()
                            .setAddSheet(new AddSheetRequest()
                                    .setProperties(new SheetProperties()
                                            .setTitle(sheet))));
                });
                BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest().setRequests(sheets);
                Spreadsheet spreadsheet = sheetService
                        .spreadsheets()
                        .create(workbook)
                        .execute();
                sheetService
                        .spreadsheets()
                        .batchUpdate(spreadsheet
                                .getSpreadsheetId(), batchUpdateRequest)
                        .execute();
                workBookId = spreadsheet.getSpreadsheetId();
                Optional<Sheet> sheetOptional = sheetService
                        .spreadsheets()
                        .get(workBookId)
                        .execute()
                        .getSheets()
                        .stream()
                        .filter(sheet -> sheet
                                .getProperties()
                                .getTitle().equals("Sheet1"))
                        .findFirst();
                if (sheetOptional.isPresent()) {
                    DeleteSheetRequest deleteSheetRequest = new DeleteSheetRequest();
                    deleteSheetRequest.setSheetId(sheetOptional.get().getProperties().getSheetId());
                    Request deleteRequest = new Request();
                    deleteRequest.setDeleteSheet(deleteSheetRequest);
                    BatchUpdateSpreadsheetRequest batchDeleteRequest = new BatchUpdateSpreadsheetRequest();
                    batchDeleteRequest.setRequests(List.of(deleteRequest));
                    sheetService.
                            spreadsheets().
                            batchUpdate(workBookId, batchDeleteRequest)
                            .execute();
                }
            }
            else
                logger.error("sheet service is null.");
        }
        catch (Exception e) {
            logger.error("Exception occurred in createWorkBookWithDefaultSheet " + e.getMessage(), e);
        }
        logger.debug("Got workbook with id " + workBookId);
        return workBookId;
    }

    public List<List<Object>> readRange(String workBookId, String sheetName, String range) {
        logger.debug("Inside method readRange for range " + range);
        List<List<Object>> values = new ArrayList<>();
        try {
            ValueRange valueRange;
            Sheets sheetService = initializeSheetService();
            String actualRange = sheetName + "!" + range;
            if (sheetService != null) {
                valueRange = sheetService
                        .spreadsheets()
                        .values()
                        .get(workBookId, actualRange)
                        .execute();
                if (valueRange != null && valueRange.getValues() != null) {
                    logger.debug("Number of rows read : " + valueRange.getValues().size());
                    values = valueRange.getValues();
                }
                else
                    logger.debug("No data found to be read for range " + range);
            }
            else
                logger.debug("Failed to read from workbook as sheetService is null");
        }
        catch (Exception e) {
            logger.error("Failed to read range " + range + " from workbook with id " + workBookId + e.getMessage(), e);
        }
        return values;
    }

    public UpdateValuesResponse writeToRange(String workBookId, String sheetName, String startingColumn, String inputType, List<List<Object>> values) {
        logger.debug("Inside method writeToRange");
        UpdateValuesResponse response = null;
        try {
            Sheets sheetService = initializeSheetService();
            String actualRange = sheetName + "!" + deriveRange(values, startingColumn);
            if (sheetService != null) {
                ValueRange valueRange = new ValueRange()
                        .setValues(values);
                response = sheetService
                        .spreadsheets()
                        .values()
                        .update(workBookId, actualRange, valueRange)
                        .setValueInputOption(inputType)
                        .execute();
                if (response != null)
                    logger.debug("Number of rows updated " + response.getUpdatedCells());
            }
            else
                logger.debug("Failed to write data to workbook as sheetService is null");

        }
        catch (Exception e) {
            logger.error("Failed to write data to range starting from " + startingColumn + " in workbook with id " + workBookId + " with input type " + inputType + " " + e.getMessage(), e);
        }
        return response;
    }

    private String deriveRange(List<List<Object>> values, String startingCell) {
        logger.debug("Deriving range with starting cell " + startingCell);
        String range = "";
        try {
            if (Boolean.FALSE.equals(values.isEmpty())) {
                int numRows = values.size();
                int numCols = values.get(0).size();
                String endCol = getColumnName(numCols + (startingCell.charAt(0) - 'A' + 1) - 1);
                int endRow = Integer.parseInt(startingCell.substring(1)) + numRows - 1;
                range =  startingCell + ":" + endCol + endRow;
            }
        }
        catch (Exception e) {
            logger.error("Exception occurred in deriving range " + e.getMessage(), e);
        }
        logger.debug("Got range " + range);
        return range;
    }

    private String getColumnName(int column) {
        logger.debug("Inside method getColumnName");
        StringBuilder columnName = new StringBuilder();
        try {
            while (column > 0) {
                int remainder = (column - 1) % 26;
                columnName.insert(0, (char) ('A' + remainder));
                column = (column - remainder) / 26;
            }
        }
        catch (Exception e) {
            logger.error("Exception occurred in getColumnName method " + e.getMessage(), e);
        }
        logger.debug("Got column name " + columnName);
        return columnName.toString();
    }

    public String findWorkBookWithName(String sheetName) {
        logger.debug("Trying to find spreadsheet with name " + sheetName);
        String workBookId = "";
        try {
            Drive driveService = initializeDriveService();
            if (driveService != null) {
                String query = "mimeType='application/vnd.google-apps.spreadsheet' and name='" + sheetName + "'";
                FileList result = driveService.files().list()
                        .setQ(query)
                        .setSpaces("drive")
                        .setFields("files(id, name)")
                        .execute();
                List<File> files = result.getFiles();
                if (files != null && !files.isEmpty()) {
                    workBookId = files.get(0).getId();
                }
            }
            else
                logger.debug("drive service is null.");
        }
        catch (Exception e) {
            logger.error("Exception occurred in finding spreadsheet by name " + e.getMessage(), e);
        }
        return workBookId;
    }

    public void addSheetToWorkBook(String workBokId, String sheetName) {
        List<String> sheetNames = new ArrayList<>();
        sheetNames.add(sheetName);
        addSheetToWorkBook(workBokId, sheetNames);
    }

    public void addSheetToWorkBook(String workBokId, List<String> sheetNames) {
        logger.debug("Inside method addSheetToWorkBook");
        try {
            Sheets sheetService = initializeSheetService();
            if (sheetService != null) {
                final List<Request> addRequests = new ArrayList<>();
                sheetNames.forEach(sheetName -> {
                    SheetProperties sheetProperties = new SheetProperties()
                            .setTitle(sheetName);
                    AddSheetRequest addSheetRequest = new AddSheetRequest()
                            .setProperties(sheetProperties);
                    Request sheetRequest = new Request()
                            .setAddSheet(addSheetRequest);
                    addRequests.add(sheetRequest);
                    logger.debug("Added request for sheet " + sheetName);
                });
                BatchUpdateSpreadsheetRequest batchUpdateAddRequest = new BatchUpdateSpreadsheetRequest()
                        .setRequests(addRequests);
                sheetService
                        .spreadsheets()
                        .batchUpdate(workBokId, batchUpdateAddRequest)
                        .execute();
                logger.debug("Added sheets " + sheetNames);
            }
            else
                logger.debug("Sheet service is null");
        }
        catch (Exception e) {
            logger.error("Exception occurred in method addSheetToWorkBook " + e.getMessage(), e);
        }
    }

    public List<String> getSheetNamesFromWorkbook(String workbookId) {
        logger.debug("Inside method getSheetNamesFromWorkbook");
        List<String> sheetNames = new ArrayList<>();
        try {
            Sheets sheetService = initializeSheetService();
            if (sheetService != null) {
                Spreadsheet spreadsheet = sheetService.
                        spreadsheets()
                        .get(workbookId)
                        .execute();
                List<Sheet> sheets = spreadsheet.getSheets();
                sheets.forEach(sheet -> sheetNames.add(sheet
                        .getProperties()
                        .getTitle()));
            }
            else
                logger.debug("sheetService is null");
        }
        catch (Exception e) {
            logger.error("Exception occurred in getSheetNamesFromWorkbook method " + e.getMessage(), e);
        }
        logger.debug("Got sheets " + sheetNames);
        return sheetNames;
    }

    public void clearSheetData(String workBookId, String sheetName) {
        List<String> sheetNames = new ArrayList<>();
        sheetNames.add(sheetName);
        clearSheetData(workBookId, sheetNames);
    }

    public void clearSheetData(String workBookId, List<String> sheetNames) {
        logger.debug("Inside method clearSheetData");
        try {
            Sheets sheetService = initializeSheetService();
            if (sheetService != null) {
                sheetNames.forEach(sheetName -> {
                    ClearValuesRequest clearRequest = new ClearValuesRequest();
                    try {
                        sheetService
                                .spreadsheets()
                                .values()
                                .clear(workBookId, sheetName, clearRequest)
                                .execute();
                    } catch (Exception e) {
                        logger.error("Exception occurred in clearing sheet " + sheetName + " " + e.getMessage(), e);
                    }
                });
            }
            else
                logger.debug("Got null sheet service");
        }
        catch (Exception e) {
            logger.error("Exception occurred in clearSheetData method " + e.getMessage(), e);
        }
    }

    public boolean isSheetExists(String workbookId, String sheetName) {
        logger.debug("Checking if " + sheetName + " exists in " + workbookId);
        boolean isSheetExists = Boolean.FALSE;
        try {
            List<String> sheetNames = getSheetNamesFromWorkbook(workbookId);
            if (sheetNames.contains(sheetName))
                isSheetExists = Boolean.TRUE;
        }
        catch (Exception e) {
            logger.error("Exception occurred in isSheetExists method " + e.getMessage(), e);
        }
        logger.debug("is sheet existing ? " + isSheetExists);
        return isSheetExists;
    }

    public void createOrClearSheets(String workbookId, String sheetName) {
        List<String> sheetNames = new ArrayList<>();
        sheetNames.add(sheetName);
        createOrClearSheets(workbookId, sheetNames);
    }

    public void createOrClearSheets(String workbookId, List<String> sheetNames) {
        logger.debug("Inside method createOrClearSheet");
        try {
            sheetNames.forEach(sheetName -> {
                if (Boolean.TRUE.equals(isSheetExists(workbookId, sheetName)))
                    clearSheetData(workbookId, sheetName);
                else
                    addSheetToWorkBook(workbookId, sheetName);
            });
        }
        catch (Exception e) {
            logger.error("Exception occurred in createOrClearSheet method " + e.getMessage(), e);
        }
    }
}
