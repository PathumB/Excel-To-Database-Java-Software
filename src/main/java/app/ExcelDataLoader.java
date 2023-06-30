package app;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;

public class ExcelDataLoader {
    private static final String FILE_PATH = "src/main/resources/health-career-me-structure.xlsx";
    private static final int SHEET_INDEX = 0;

    public static void loadDataFromExcel() {
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
                String name = row.getCell(2).getStringCellValue();
                String email = row.getCell(5).getStringCellValue();
                String gender = row.getCell(13).getStringCellValue();
                String password = row.getCell(22).getStringCellValue();
                String file1 = row.getCell(25).getStringCellValue();

                if (email.isEmpty()) {
                    System.out.println("Skipping row " + rowIndex + " - email value is empty");
                    continue;
                }

                // Insert data into the users table
                String insertUserQuery = "INSERT INTO users (name, email, password, gender, role, created_at, updated_at) " +
                        "VALUES (?, ?, ?, ?,?, NOW(), NOW())";
                PreparedStatement userStatement = connection.prepareStatement(insertUserQuery, Statement.RETURN_GENERATED_KEYS);
                userStatement.setString(1, name);
                userStatement.setString(2, email);
                userStatement.setString(3, password);
                userStatement.setString(4, gender);
                userStatement.setString(5, "candidate");
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

    public static void main(String[] args) {
        ExcelDataLoader.loadDataFromExcel();
    }
}