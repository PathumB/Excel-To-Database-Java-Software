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
    private static final String FILE_PATH = "src/main/resources/health-career-me-structure.xlsx";
    private static final int SHEET_INDEX = 0;

    public static void loadDataFromExcel() {
        String inputFormat = "dd-MM-yyyyHH:mm";
        String formatDate = "yyyy-MM-dd HH:mm:ss";

        try (FileInputStream file = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(file);
             Connection connection = DriverManager.getConnection(
                     DatabaseConfig.DB_URL, DatabaseConfig.DB_USER, DatabaseConfig.DB_PASSWORD
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

                // Format dates
                try {
                    approved_at = formatDate(approved_at, inputFormat, formatDate);
                    created_at = formatDate(created_at, inputFormat, formatDate);
                } catch (ParseException e) {
                    e.printStackTrace();
                }

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

                // Insert data into the candidate_cvs table
                String insertCVQuery = "INSERT INTO candidate_cvs (user_id, file, created_at, updated_at) " +
                        "VALUES (?, ?, NOW(), NOW())";
                PreparedStatement cvStatement = connection.prepareStatement(insertCVQuery);
                cvStatement.setInt(1, userId);
                cvStatement.setString(2, file1);
                cvStatement.executeUpdate();

                System.out.println("Inserted data for row: " + rowIndex);
            }
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }

    // date format
    private static String formatDate(String dateStr, String inputFormat, String outputFormat) throws ParseException {
        SimpleDateFormat inputDateFormat = new SimpleDateFormat(inputFormat);
        SimpleDateFormat outputDateFormat = new SimpleDateFormat(outputFormat);

        Date date = inputDateFormat.parse(dateStr);
        return outputDateFormat.format(date);
    }

    public static void main(String[] args) {
        ExcelDataLoader.loadDataFromExcel();
    }
}