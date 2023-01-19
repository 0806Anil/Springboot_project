package io.xlsxtojson.demo;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@RestController

public class ExcelToJSONConverter {
	
	@PostMapping("/convert")
    public JSONArray convertXlsxToJson(@RequestParam("file") MultipartFile file) {
        JSONArray jsonArray = new JSONArray();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            Row row;
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                List<String> data = new ArrayList<>();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            data.add(String.valueOf(cell.getNumericCellValue()));
                            break;
                        case STRING:
                            data.add(cell.getStringCellValue());
                            break;
                    }
                }
                JSONObject json = new JSONObject();
                for (int i = 0; i < data.size(); i++) {
                    json.put("column" + (i + 1), data.get(i));
                }
                jsonArray.add(json);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return jsonArray;
    }
		
    
}





