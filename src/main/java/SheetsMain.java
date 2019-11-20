import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.FileContent;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;

public class SheetsMain {
    private static Credential authCredential;
    private static Sheets sheetsService;
    private static Drive driveService;
    private static String APPLICATION_NAME = "Sheets_project";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    private static Credential autorize() throws IOException, GeneralSecurityException {
        InputStream inputStream = SheetsMain.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(inputStream));

        List<String> scopes = new ArrayList<String>();
        scopes.add(SheetsScopes.DRIVE);
        scopes.add(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
                .setAccessType("offline")
                .build();

        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");

        return credential;
    }

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, authCredential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static Drive getDriveService() throws IOException, GeneralSecurityException {
        return new Drive.Builder(GoogleNetHttpTransport.newTrustedTransport(), JSON_FACTORY, authCredential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static String uploadCSVfile(java.io.File javaFile) throws IOException {
        File fileMetadata = new File();
        fileMetadata.setName(javaFile.getName());
        fileMetadata.setMimeType("application/vnd.google-apps.spreadsheet");

        FileContent mediaContent = new FileContent("text/csv", javaFile);
        File file = driveService.files().create(fileMetadata, mediaContent)
                .setFields("id")
                .execute();
        System.out.println(javaFile.getName());

        return "https://docs.google.com/spreadsheets/d/" + file.getId();
    }

    public static List<List<Object>> getData(String spreadsheetId, String range) throws IOException {
    return sheetsService.spreadsheets().values().get(spreadsheetId, range).execute().getValues();
    }

    public static void setBGcolorRed(String spreadsheetId, Integer gid, Integer[] gridRange) throws IOException {
        List<Request> requests = new ArrayList();

        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setCell(new CellData()
                                .setUserEnteredFormat(new CellFormat()
                                        .setBackgroundColor(new Color()
                                                .setRed(Float.valueOf("1"))
                                                .setGreen(Float.valueOf("0"))
                                                .setBlue(Float.valueOf("0"))
                                        )
                                )
                        )
                        .setRange(new GridRange()
                                .setSheetId(gid)
                                .setStartRowIndex(gridRange[0])
                                .setEndRowIndex(gridRange[1])
                                .setStartColumnIndex(gridRange[2])
                                .setEndColumnIndex(gridRange[3])
                        )
                        .setFields("userEnteredFormat(backgroundColor)")
                )
        );

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
    }

    public static void main(String[] args) throws IOException, GeneralSecurityException {
        authCredential = autorize();
        sheetsService = getSheetsService();
        driveService = getDriveService();

//        java.io.File file = new java.io.File("src/main/resources/sheet.csv");
//        String link = uploadCSVfile(file);
//        System.out.println("Google spreadsheets link: " + link);
//
//        String range = "A1:E2";
//
//        String[] split = link.split("/");
//        String sheetId = split[split.length-1];
//        List<List<Object>> data = getData(sheetId, range);
//        System.out.println(data);

        String spreadsheetId = "1XkQHVFPl52aAnP8TKP2DVRpZe5Odo3NujSXyGXuv7aY";
        Integer gid = 2008177436;
        Integer[] gridRange = {0,2,0,5};
        setBGcolorRed(spreadsheetId, gid, gridRange);
    }
}
