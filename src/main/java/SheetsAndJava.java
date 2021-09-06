import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

public class SheetsAndJava {

    private static Sheets sheetsService;
    private static String APPLICATION_NAME = "Google Sheets Tunts Test";
    private static String SPREADSHEET_ID = "1aeRs9cpnG_nUiAaPSYCeHMq0jwjqR8PbA26X_Eyik4o";

    private static Credential authorize() throws IOException, GeneralSecurityException {
        InputStream in = SheetsAndJava.class.getResourceAsStream("/credentials.json");


        System.out.println(in);
        assert in != null;
        JacksonFactory jacksonFactory = new JacksonFactory();
        HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(
                jacksonFactory, new InputStreamReader(in)
        );

        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(httpTransport,
                jacksonFactory,
                clientSecrets,
                scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
                .build();

        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");

        return credential;
    }

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        Credential credential = authorize();

        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(), credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static long getMedium(Object p1, Object p2, Object p3) {
        double result = ((Double) p1 +(Double) p2 +(Double) p3) / 3;
        return Math.round(result);
    }

    public static long getFinalGrade(double medium) {
        double gradeNeeded = 100 - medium;
        return Math.round(gradeNeeded * 2) ;
    }

    public static void setSituationAndGrade(String situation, Integer gradeToApproved  ) throws IOException {
        String situationRange = "engenharia_de_software!G4:G27";
        String toPassRange = "engenharia_de_software!H4:H27";

        ValueRange situationValue = new ValueRange().setValues(Arrays.asList(Arrays.asList(situation)));
        UpdateValuesResponse updateSituation = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, situationRange, situationValue)
                .setValueInputOption("RAW")
                .execute();


        ValueRange toPassValue = new ValueRange().setValues(Arrays.asList(Arrays.asList(gradeToApproved)));
        UpdateValuesResponse updateToPass = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, toPassRange, toPassValue)
                .setValueInputOption("RAW")
                .execute();
    }

    public static void main(String[] args) throws IOException, GeneralSecurityException {
        sheetsService = getSheetsService();

        String range = "engenharia_de_software!A4:F27";

        int missesLimit = 15;

        ValueRange sheet = sheetsService.spreadsheets().values().get(SPREADSHEET_ID, range).execute();
        List<List<Object>> values = sheet.getValues();


        if (values == null || values.isEmpty()) {
            System.out.println("No data");
        } else {
            for (List row : values) {
                System.out.println("Getting students");
                System.out.printf("number: %s , student: %s , misses: %s, p1: %s, p2: %s, p3: %s \n", row.get(0), row.get(1), row.get(2), row.get(3), row.get(4), row.get(5));
                System.out.println("verifying student misses");
                double studentMisses = (double) row.get(2);
                if (studentMisses > missesLimit) {
                    setSituationAndGrade("Reprovado por Falta" ,  0);
                    continue;
                }

                double studentMedium = getMedium(row.get(3), row.get(4), row.get(5));

                System.out.println("Verifying student grades");
                if (studentMedium < 50) {
                    setSituationAndGrade("Reprovado por Nota", 0);
                } else if (studentMedium < 70) {
                    long gradeNeeded = getFinalGrade(studentMedium);
                    setSituationAndGrade("Exame Final", (int) gradeNeeded);
                } else {
                    setSituationAndGrade("Aprovado", 0);
                }
            }

            System.out.println("Finished setting grades");

        }
    }


}
