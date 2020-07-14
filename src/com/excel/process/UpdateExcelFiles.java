package com.excel.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * This program is used to read an excel file and check if the numbers
 * within it exist in the database and accordingly marks them in green color.
 * Once the program is completed, user can open the excel file and filter
 * it by the green color to ascertain the existing numbers
 */

public class UpdateExcelFiles {

	static int countOfFilesProcessed = 0;

	static UpdateExcelFiles obj = new UpdateExcelFiles();

	static Connection conn = null;
	static Statement stmt = null;
	static ResultSet rs = null;

	public Connection getConnection() {

		try {
			conn = DriverManager
					.getConnection("jdbc:sqlserver://localhost:1433;databaseName=students;integratedSecurity=true;");
		} catch (SQLException e) {
			e.printStackTrace();
		}

		return conn;
	}

	public Statement getStatement() {
		try {
			stmt = conn.createStatement();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return stmt;
	}

	public long returnAValidNumber(long phone_Number) {
		long phoneNumber = 0L;

		if (phone_Number > 2010000000L && phone_Number <= 9999999999L) {
			phoneNumber = phone_Number;
		}
		if (phone_Number >= 10000000000L && phone_Number <= 19999999999L) {
			if (((phone_Number % 10000000000L) > 2010000000L)) {
				phoneNumber = phone_Number % 10000000000L;
			}
		}
		return phoneNumber;
	}

	public static void processExcelFile(File file) {

		long value = 0L;

		Workbook workbook = null;
		FileInputStream excelInputStream = null;
		FileOutputStream fileOut = null;

		try {
			excelInputStream = new FileInputStream(file.getCanonicalPath());
		} catch (IOException e) {
			e.printStackTrace();
		}

		try {
			if (file.getCanonicalPath().toLowerCase().endsWith(".xls")) {
				workbook = new HSSFWorkbook(excelInputStream);
				System.out.println("Reading file: " + file.getName());
			} else if (file.getCanonicalPath().toLowerCase().endsWith(".xlsx")) {
				workbook = new XSSFWorkbook(excelInputStream);
				System.out.println("Reading file: " + file.getName());
			}
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		for (Sheet sheet : workbook) {
			for (Row row : sheet) {
				for (Cell cell : row) {
					switch (cell.getCellType()) {

					case STRING:

						Pattern pattern = Pattern.compile("[a-zA-Z]");
						Matcher matcher = pattern.matcher(cell.getStringCellValue());

						if (!matcher.find()) {
							try {
								value = Long.parseLong(cell.getStringCellValue().replaceAll("[^0-9]", ""));
							} catch (NumberFormatException n) {
								value = 0L;
							}
						}

						value = obj.returnAValidNumber(value);

						if (value != 0) {
							String sql = "select PhoneNumber from ContactInformation where PhoneNumber = " + value;

							try {
								rs = stmt.executeQuery(sql);

								while (rs.next()) {
									cell.setCellStyle(style);
								}
							} catch (SQLException e) {
								e.printStackTrace();
							}
						}

						value = 0L;
						break;

					case NUMERIC:

						value = (long) cell.getNumericCellValue();
						value = obj.returnAValidNumber(value);

						if (value != 0) {
							String sql = "select PhoneNumber from ContactInformation where PhoneNumber = " + value;

							try {
								rs = stmt.executeQuery(sql);

								while (rs.next()) {
									cell.setCellStyle(style);
								}
							} catch (SQLException e) {
								e.printStackTrace();
							}
						}

					case _NONE:
						break;
					case BOOLEAN:
						break;
					case FORMULA:
						break;
					case BLANK:
						break;
					case ERROR:
						break;

					default:
						break;
					}
				}
			}
		}

		try {
			excelInputStream.close();
			fileOut = new FileOutputStream(new File(file.getCanonicalPath()));
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws IOException, InterruptedException {
		long start = System.currentTimeMillis();
		conn = obj.getConnection();
		stmt = obj.getStatement();
		
		Scanner scanner = new Scanner(System.in);
		System.out.print("Enter the folder you want to search: ");
		String inputFolder = scanner.nextLine();
		

		File currentDir = new File(inputFolder);
		System.out.println("Searching for Excel files in: " + currentDir.getCanonicalPath());
		System.out.println("===========================================");
		createListOfExcelFiles(currentDir);
		scanner.close();
		System.out.println("===========================================");
		System.out.println("Process complete");
		System.out.println("===========================================");
		System.out.println("Files found: " + countOfFilesProcessed);
		long end = System.currentTimeMillis();

		System.out.println("Total time required: " + ((end - start) / 60000) + " in minutes");
		System.out.println("Total time required: " + (end - start) + " in milliseconds");
		System.out.println("===========================================");
		
		try {
			if (rs != null && !rs.isClosed()) {
				rs.close();
			}

			if (stmt != null && !stmt.isClosed()) {
				stmt.close();
			}

			if (conn != null && !conn.isClosed()) {
				conn.close();
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}

	public static void createListOfExcelFiles(File dir) throws IOException, InterruptedException {

		try {
			File[] files = dir.listFiles();
			for (File file : files) {

				if (file.isDirectory()) {
					createListOfExcelFiles(file);
				} else if ((file.getName().toLowerCase().endsWith(".xls")
						|| (file.getName().toLowerCase().endsWith(".xlsx")))
						&& !file.getName().toLowerCase().startsWith("~$")) {
					processExcelFile(file);
					countOfFilesProcessed++;
				}
			}
		} catch (IOException i) {
			i.printStackTrace();
		} catch (NullPointerException n) {
			n.printStackTrace();
		}
	}

}
