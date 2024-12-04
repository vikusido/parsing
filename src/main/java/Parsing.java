//use maven and pom, write in Excel format, 22 version of Java!!!

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;

public class Parsing {

    public static void main(String[] args) throws IOException {

        String[] apiUrls = {"https://fake-json-api.mock.beeceptor.com/users", "https://fake-json-api.mock.beeceptor.com/companies"};
        String[] excelFiles = {"users.xlsx", "companies.xlsx"};

        ObjectMapper mapper = new ObjectMapper();
        for (int i = 0; i < apiUrls.length; i++) {
            String apiUrl = apiUrls[i];
            String excelFile = excelFiles[i];
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("API Data");
            int rowIndex = 0;
            String jsonData = fetchJsonData(apiUrl);
            // Parse JSON data and write it to the Excel sheet.  Error handling included.
            try {
                JsonNode rootNode = mapper.readTree(jsonData);
                writeDataToSheet(sheet, rootNode, rowIndex);
            } catch (IOException e) {
                System.err.println("Error parsing JSON from " + apiUrl + ": " + e.getMessage());
            }

            try (FileOutputStream outputStream = new FileOutputStream(excelFile)) {
                workbook.write(outputStream);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            System.out.println("Data is written to " + excelFile);
            workbook.close();
        }
    }

    // Writes the header row and data rows to the Excel sheet.
    private static void writeDataToSheet(Sheet sheet, JsonNode rootNode, int rowIndex) {
        Row headerRow = sheet.createRow(rowIndex++);
        headerRow.createCell(0).setCellValue("Index");
        headerRow.createCell(1).setCellValue("ID");
        headerRow.createCell(2).setCellValue("Name");
        headerRow.createCell(3).setCellValue("Email");
        headerRow.createCell(4).setCellValue("Industry");

        if (rootNode.isArray()) {
            for (int i = 0; i < rootNode.size(); i++) {
                JsonNode item = rootNode.get(i);
                writeJsonNodeToRow(sheet, item, rowIndex++, i + 1);

            }
        } else {
            writeJsonNodeToRow(sheet, rootNode, rowIndex++, 1);
        }
    }

    // Writes a single JSON object's data to a row in the Excel sheet.
    private static void writeJsonNodeToRow(Sheet sheet, JsonNode jsonNode, int rowIndex, int index) {
        Row row = sheet.createRow(rowIndex);
        int cellIndex = 0;
        row.createCell(cellIndex++).setCellValue(index); //Index
        //using ternary operator
        row.createCell(cellIndex++).setCellValue(jsonNode.has("id") ? jsonNode.get("id").asInt() : 0); //ID
        row.createCell(cellIndex++).setCellValue(jsonNode.has("name") ? jsonNode.get("name").asText() : ""); //Name
        row.createCell(cellIndex++).setCellValue(jsonNode.has("email") ? jsonNode.get("email").asText() : ""); //Email
        row.createCell(cellIndex++).setCellValue(jsonNode.has("industry") ? jsonNode.get("industry").asText() : ""); //Industry
    }


    // Fetches JSON data from a given API URL.
    private static String fetchJsonData(String apiUrl) throws IOException {
        URL url = new URL(apiUrl);
        URLConnection connection = url.openConnection();
        BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()));
        StringBuilder response = new StringBuilder();
        String line;
        while ((line = reader.readLine()) != null) {
            response.append(line);
        }
        reader.close();
        return response.toString();
    }
}
