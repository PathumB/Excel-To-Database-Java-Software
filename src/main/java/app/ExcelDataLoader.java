package app;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.common.hash.Hashing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mindrot.jbcrypt.BCrypt;
import software.amazon.awssdk.auth.credentials.DefaultCredentialsProvider;
import software.amazon.awssdk.core.sync.RequestBody;
import software.amazon.awssdk.regions.Region;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.s3.model.PutObjectRequest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.security.GeneralSecurityException;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

// 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹main method🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
public class ExcelDataLoader {
    public static void main(String[] args) {
        new LoadData().loadDataFromExcel();
    }
}

// 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹whole process🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
class LoadData {
    private static final String FILE_PATH = "src/main/resources/health-career-me-structure.xlsx";
    private static final int SHEET_INDEX = 0;
    private static final String SERVICE_ACCOUNT_JSON_PATH = "src/main/resources/project-02-391607-f0109bc4ffcd.json";


    public static void loadDataFromExcel() {
        String inputFormat = "dd-MM-yyyyHH:mm";
        String formatDate = "yyyy-MM-dd HH:mm:ss";

        // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸DB connection🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
        try (FileInputStream file = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(file);
             Connection connection = DriverManager.getConnection(
                     env.DB_URL, env.DB_USER, env.DB_PASSWORD
             ))
        {

            Sheet sheet = workbook.getSheetAt(SHEET_INDEX);

            // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸extract data from excel & insert to DB🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
            // Start from row 1 to skip the header
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Get data from excel🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                int id = (int) row.getCell(0).getNumericCellValue();
                String name = row.getCell(1).getStringCellValue();
                String email = row.getCell(2).getStringCellValue();
                String gender = row.getCell(3).getStringCellValue();
                String password = row.getCell(4).getStringCellValue();
                String file1 = row.getCell(5).getStringCellValue();
                String approved_at = row.getCell(6).getStringCellValue();
                String created_at = row.getCell(7).getStringCellValue();
                String licenseTypes = row.getCell(9).getStringCellValue();


                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Format dates🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                try {
                    approved_at = formatDate(approved_at, inputFormat, formatDate);
                    created_at = formatDate(created_at, inputFormat, formatDate);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸validate emails🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                if (email.isEmpty()) {
                    System.out.println("Skipping row " + rowIndex + " - email value is empty");
                    continue;
                }

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸hash password🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                String hashedPassword = BCrypt.hashpw(password, BCrypt.gensalt());

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Insert data into the users table🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                String insertUserQuery = "INSERT INTO users (name, email, password, gender, role, approved_at, created_at, updated_at) " +
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
                PreparedStatement userStatement = connection.prepareStatement(insertUserQuery, Statement.RETURN_GENERATED_KEYS);
                userStatement.setString(1, name);
                userStatement.setString(2, email);
                userStatement.setString(3, hashedPassword);
                userStatement.setString(4, gender);
                userStatement.setString(5, "candidate");
                userStatement.setString(6, approved_at);
                userStatement.setString(7, created_at);
                userStatement.setString(8, created_at);
                userStatement.executeUpdate();

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Get the auto-generated user_id🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                ResultSet generatedKeys = userStatement.getGeneratedKeys();
                int userId = 0;
                if (generatedKeys.next()) {
                    userId = generatedKeys.getInt(1);
                }

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Insert license types🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                String[] licenseTypeArray = licenseTypes.split(",");
                for (String licenseType : licenseTypeArray) {
                    insertCandidateLicense(connection, userId, licenseType.trim(), created_at);
                }

                // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Download file from Google Drive (method call)🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
                String[] links = file1.split("\\s*,\\s*");
                for (String link : links) {
                    String fileId = extractFileId(link.trim());
                    if (fileId != null) {
                        try {
                            downloadFileFromGoogleDrive(fileId, connection, userId);
                        }catch (IOException | GeneralSecurityException | org.apache.http.ParseException e) {
                            e.printStackTrace();
                        }
                    } else {
                        System.out.println("Invalid or unsupported Google Drive link: " + link);
                    }
                }

                System.out.println("Inserted data for row: " + rowIndex);
            }
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }


    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹Extract file ID from Google Drive link🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    public static String extractFileId(String googleDriveLink) {
        // Define the regex pattern to match the file ID in the Google Drive link
        Pattern pattern = Pattern.compile("[-\\w]{25,}");
        Matcher matcher = pattern.matcher(googleDriveLink);

        // Find the first occurrence of the regex pattern in the link
        if (matcher.find()) {
            return matcher.group();
        }

        // If no match found, return null or throw an exception as needed
        return null;
    }

    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹download file from Google Drive🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    public static void downloadFileFromGoogleDrive(String fileId,Connection connection, int userId) throws IOException, GeneralSecurityException {
        HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
        JsonFactory jsonFactory = GsonFactory.getDefaultInstance();

        GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(SERVICE_ACCOUNT_JSON_PATH))
                .createScoped(Collections.singleton(DriveScopes.DRIVE_READONLY));

        Drive drive = new Drive.Builder(httpTransport, jsonFactory, setHttpTimeout(credential))
                .setApplicationName(env.APPLICATION_NAME)
                .setHttpRequestInitializer(credential)
                .build();

        // Get the file metadata
        File fileMetadata = drive.files().get(fileId).execute();
        // set hash name
        String hashedFileName = null;
        try {
            hashedFileName = hashFileName(fileMetadata.getName());
        }catch (ParseException e){
            e.printStackTrace();
        }

        // set file path
        String localFilePath = "src/main/resources/cvs/" + hashedFileName;

        try (OutputStream outputStream = new FileOutputStream(localFilePath)) {
            drive.files().get(fileId).executeMediaAndDownloadTo(outputStream);
        }

        System.out.println("File downloaded: " + localFilePath);
        try {
            // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸insert data to cvs table🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
            insertCandidateCV(connection, userId, hashedFileName);

            // 🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸Upload the file to Amazon S3🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸🔸
            String s3BucketName = env.S3_BUCKET_NAME;
            String s3Folder = "storage/app/cvs";
            String s3Key = s3Folder + "/" + Paths.get(localFilePath).getFileName();
            uploadFileToS3(localFilePath, s3BucketName, s3Key);

        }catch (SQLException e){
            e.printStackTrace();
        }
    }

    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹Upload the file to Amazon S3 bucket🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    private static void uploadFileToS3(String localFilePath, String bucketName, String s3Key) {
        S3Client s3Client = S3Client.builder()
                .region(Region.US_EAST_1)
                .credentialsProvider(DefaultCredentialsProvider.create())
                .build();

        try {
            // Read the file as bytes from the local file path
            byte[] fileBytes = Files.readAllBytes(Paths.get(localFilePath));

            // Upload the file to Amazon S3
            s3Client.putObject(PutObjectRequest.builder()
                            .bucket(bucketName)
                            .key(s3Key)
                            .build(),
                    RequestBody.fromBytes(fileBytes));
            System.out.println("File uploaded to Amazon S3: " + s3Key);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            s3Client.close();
        }
    }

    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹set timeout🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    private static HttpRequestInitializer setHttpTimeout(final HttpRequestInitializer requestInitializer) {
        return httpRequest -> {
            requestInitializer.initialize(httpRequest);
            httpRequest.setReadTimeout(3 * 60000); // 3 minutes
        };
    }

    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹hash file name🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    private static String hashFileName(String fileName) throws ParseException {
        String[] fileNameArray = fileName.split("\\.");
        String extension = fileNameArray[fileNameArray.length - 1];
        String fileNameWithoutExtension = fileName.substring(0, fileName.length() - extension.length() - 1);
        String hashedFileName = Hashing.sha256().hashString(fileNameWithoutExtension, StandardCharsets.UTF_8).toString();
        return hashedFileName + "." + extension;
    }


    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹insert into candidate_cvs table🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    private static void insertCandidateCV(Connection connection, int userId, String fileName) throws SQLException{
        String insertCVQuery = "INSERT INTO candidate_cvs (user_id, file, created_at, updated_at) " +
                "VALUES (?, ?, NOW(), NOW())";
        PreparedStatement cvStatement = connection.prepareStatement(insertCVQuery);
        cvStatement.setInt(1, userId);
        cvStatement.setString(2, fileName);
        cvStatement.executeUpdate();
    }

    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹Insert into candidate_licenses table🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    private static void insertCandidateLicense(Connection connection, int userId, String licenseType, String createdAt) throws SQLException {
        String insertLicenseQuery = "INSERT INTO candidate_licenses (user_id, title, created_at, updated_at) " +
                "VALUES (?, ?, ?, ?)";

        PreparedStatement licenseStatement = connection.prepareStatement(insertLicenseQuery);
        licenseStatement.setInt(1, userId);
        licenseStatement.setString(2, licenseType);
        licenseStatement.setString(3, createdAt);
        licenseStatement.setString(4, createdAt);

        licenseStatement.executeUpdate();
    }

    // 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹date format🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
    private static String formatDate(String dateStr, String inputFormat, String outputFormat) throws ParseException {
        SimpleDateFormat inputDateFormat = new SimpleDateFormat(inputFormat);
        SimpleDateFormat outputDateFormat = new SimpleDateFormat(outputFormat);

        Date date = inputDateFormat.parse(dateStr);
        return outputDateFormat.format(date);
    }
}
