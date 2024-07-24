package com.mainfolder.Controller;

import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;

import java.io.*;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

@RestController
public class MainController {

	@PostMapping("/getData")
	public ResponseEntity<String> changeXLData() throws IOException {
		String inputFile = "D:\\Bin\\Cash Statment - 1.BBKOPR CP.xlsx";
		String outputFile = "D:\\Bin\\output.csv";
		try {
			InputStream inputStream = new FileInputStream(inputFile);
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			Map<String, String> headerInfo = extractHeaderInfo(sheet);
			// header Info EG: {Customer=1000567, Currency=BHD, Account No=12345678}
			List<String[]> data = processSheet(sheet, headerInfo);
			saveAsCSV(data, outputFile);
			for (String[] a : data) {
				System.out.println(Arrays.toString(a));
			}
			workbook.close();
		} catch (Exception e) {
			System.out.println(e);
		}
		return ResponseEntity.status(HttpStatus.OK).body("Success");
	}

	private Map<String, String> extractHeaderInfo(Sheet sheet) {
		Map<String, String> headerInfo = new HashMap<>();
		for (Row row : sheet) {
			for (Cell cell : row) {
				String cellValue = getCellValue(cell);
				if ((cellValue.contains("Account :")) || (cellValue.contains("Account Number:"))) {
					String[] parts = cellValue.split(":");
					if (parts.length > 1) {
						headerInfo.put("Account No", parts[1].trim());
					} else {
						Cell nextCell = row.getCell(cell.getColumnIndex() + 1);
						headerInfo.put("Account No", getCellValue(nextCell).trim());
					}
				} else if ((cellValue.contains("Customer :")) || (cellValue.contains("Customer:"))) {
					String[] parts = cellValue.split(":");
					if (parts.length > 1) {
						headerInfo.put("Customer", parts[1].trim());
					} else {
						Cell nextCell = row.getCell(cell.getColumnIndex() + 1);
						headerInfo.put("Customer", getCellValue(nextCell).trim());
					}
				} else if ((cellValue.contains("Currency :")) || (cellValue.contains("Currency:"))) {
					String[] parts = cellValue.split(":");
					if (parts.length > 1) {
						headerInfo.put("Currency", parts[1].trim());
					} else {
						Cell nextCell = row.getCell(cell.getColumnIndex() + 1);
						headerInfo.put("Currency", getCellValue(nextCell).trim());
					}
				}
			}
		}
		return headerInfo;
	}

	private List<String[]> processSheet(Sheet sheet, Map<String, String> headerInfo) {
		List<String[]> data = new ArrayList<>();
		Iterator<Row> rowIterator = sheet.iterator();
		boolean headerFound = false;
		List<String> headerData = new ArrayList<>();

		while (rowIterator.hasNext() && !headerFound) {
			Row row = rowIterator.next();
			for (Cell cell : row) {
				String cellValue = getCellValue(cell);
				if (!cellValue.trim().isEmpty()) {
					headerData.add(cellValue);
				}
				if (cell.getCellType() == CellType.STRING && "Value date".equalsIgnoreCase(cell.getStringCellValue())) {
					headerFound = true;
				}
			}
			if (headerFound) {
				// Add additional headers
				headerData.add("Account No");
				headerData.add("Customer");
				headerData.add("Currency");
				data.add(headerData.toArray(new String[0]));
			} else {
				headerData.clear();
			}
		}

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			List<String> rowData = new ArrayList<>();

			for (Cell cell : row) {
				String cellValue = getCellValue(cell);
				if (!cellValue.trim().isEmpty()) {
					rowData.add(cellValue);
				}
			}

			if (rowData.size() >= (headerData.size() - 3)) {
				// Add header information (Account No, Customer, Currency)
				rowData.add(headerInfo.getOrDefault("Account No", ""));
				rowData.add(headerInfo.getOrDefault("Customer", ""));
				rowData.add(headerInfo.getOrDefault("Currency", ""));
				data.add(rowData.toArray(new String[0]));
			}
		}

		return data;
	}

	private String getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
				return sdf.format(date);
			} else {
				return BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString();
			}
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case FORMULA:
			return cell.getCellFormula();
		default:
			return "";
		}
	}

	private void saveAsCSV(List<String[]> data, String fileName) throws IOException {
		FileWriter csvWriter = new FileWriter(fileName);
		for (String[] rowData : data) {
			List<String> formattedRowData = new ArrayList<>();
			for (String cellData : rowData) {
				if (cellData.contains(",")) {
					formattedRowData.add("\"" + cellData + "\"");
				} else {
					formattedRowData.add(cellData);
				}
			}
			csvWriter.append(String.join(",", formattedRowData));
			csvWriter.append("\n");
		}
		csvWriter.flush();
		csvWriter.close();
	}
}
