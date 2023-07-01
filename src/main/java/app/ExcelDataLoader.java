package app;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mindrot.jbcrypt.BCrypt;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelDataLoader {
    public static void main(String[] args) {
        new LoadData().loadDataFromExcel();
    }
}

class LoadData {
    private static final String FILE_PATH = "src/main/resources/health-career-me-structure.xlsx";
    private static final int SHEET_INDEX = 0;

    public static void loadDataFromExcel() {
        String inputFormat = "dd-MM-yyyyHH:mm";
        String formatDate = "yyyy-MM-dd HH:mm:ss";

        try (FileInputStream file = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(file);
             Connection connection = DriverManager.getConnection(
                     env.DB_URL, env.DB_USER, env.DB_PASSWORD
             ))
        {

            Sheet sheet = workbook.getSheetAt(SHEET_INDEX);

            // Start from row 1 to skip the header
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                int id = (int) row.getCell(0).getNumericCellValue();
                String name = row.getCell(1).getStringCellValue();
                String email = row.getCell(2).getStringCellValue();
                String gender = row.getCell(3).getStringCellValue();
                String password = row.getCell(4).getStringCellValue();
                String file1 = row.getCell(5).getStringCellValue();
                String approved_at = row.getCell(6).getStringCellValue();
                String created_at = row.getCell(7).getStringCellValue();
                String licenseTypes = row.getCell(9).getStringCellValue();

                // Format dates
                try {
                    approved_at = formatDate(approved_at, inputFormat, formatDate);
                    created_at = formatDate(created_at, inputFormat, formatDate);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

                // validation
                if (email.isEmpty()) {
                    System.out.println("Skipping row " + rowIndex + " - email value is empty");
                    continue;
                }

                // hash password
                String hashedPassword = BCrypt.hashpw(password, BCrypt.gensalt());

                // Insert data into the users table
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

                // Get the auto-generated user_id
                ResultSet generatedKeys = userStatement.getGeneratedKeys();
                int userId = 0;
                if (generatedKeys.next()) {
                    userId = generatedKeys.getInt(1);
                }

                // candicate_cvs table
                insertCandidateCV(connection, userId, file1);

                // Insert license types
                String[] licenseTypeArray = licenseTypes.split(",");
                for (String licenseType : licenseTypeArray) {
                    insertCandidateLicense(connection, userId, licenseType.trim(), created_at);
                }

                System.out.println("Inserted data for row: " + rowIndex);
            }
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }


    // insert into candidate_cvs table
    private static void insertCandidateCV(Connection connection, int userId, String file1) throws SQLException{
        String insertCVQuery = "INSERT INTO candidate_cvs (user_id, file, created_at, updated_at) " +
                "VALUES (?, ?, NOW(), NOW())";
        PreparedStatement cvStatement = connection.prepareStatement(insertCVQuery);
        cvStatement.setInt(1, userId);
        cvStatement.setString(2, file1);
        cvStatement.executeUpdate();
    }

    // Insert into candidate_licenses table
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

    // date format
    private static String formatDate(String dateStr, String inputFormat, String outputFormat) throws ParseException {
        SimpleDateFormat inputDateFormat = new SimpleDateFormat(inputFormat);
        SimpleDateFormat outputDateFormat = new SimpleDateFormat(outputFormat);

        Date date = inputDateFormat.parse(dateStr);
        return outputDateFormat.format(date);
    }
}