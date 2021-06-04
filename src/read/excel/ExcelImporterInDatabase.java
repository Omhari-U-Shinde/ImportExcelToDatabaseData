package read.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.log4j.Appender;
import org.apache.log4j.FileAppender;
import org.apache.log4j.Layout;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelImporterInDatabase {
	static Logger logger = Logger.getLogger(ExcelImporterInDatabase.class.getName());
	XSSFRow row;
	private final String url = "jdbc:postgresql://localhost/postgres";
	private final String user = "postgres";
	private final String password = "Shinde@123";
	static int recordInserted = 0;
	static int recordNotInserted = 0;
	List list = new ArrayList();
	List<String> excelColumnList = new ArrayList<>();
	List<String> tableColumnList = new ArrayList<>();
	static List notInsertedRecorLlist = new ArrayList();

	boolean flag = true;
	static String[] fileNameExcel;

	public static void main(String[] args) throws ClassNotFoundException, IOException {

		Layout layout = new PatternLayout(" %d{yyyy-MM-dd--hh:mm}  %p-- %m %n");
		Appender app = new FileAppender(layout, "LogFile/data.log");
		logger.addAppender(app);

		String folderpath = "Resource\\";
		File file1 = new File(folderpath);
		String absolutePath1 = file1.getAbsolutePath();
		// logger.debug(absolutePath1);

		ExcelImporterInDatabase database = new ExcelImporterInDatabase();

		File folder = new File(absolutePath1);
		database.getAllFilesWithCertainExtension(folder, "xlsx");

		String fn = "";
		for (int j = 0; j < fileNameExcel.length; j++) {
			fn = "";
			for (int i = 0; i < fileNameExcel[j].length(); i++) {
				if (fileNameExcel[j].charAt(i) == '.') {
					break;
				} else {
					fn = fn + fileNameExcel[j].charAt(i);
				}

			}

			String url = "Resource\\" + fn + ".xlsx";// Resource
			File file = new File(url);
			if (file.exists()) {

				String absolutePath = file.getAbsolutePath();
				database.readFile(absolutePath, fn);

				if (recordInserted == 0 && recordNotInserted == 0) {
					logger.info("No Record Available in File");
				} else {
					logger.info(recordInserted + " Records  Inserted");
					logger.info(recordNotInserted + " Records Not Inserted");

					if (!notInsertedRecorLlist.isEmpty()) {

						for (int i = 0; i < notInsertedRecorLlist.size(); i++) {

							logger.warn("" + notInsertedRecorLlist.get(i));
						}
					}
				}
				notInsertedRecorLlist.clear();
				recordInserted = 0;
				recordNotInserted = 0;
			} else {
				logger.debug("File not found ");
			}
		}

	}

	public void readFile(String fileName, String tablename) {
		flag = true;
		FileInputStream fis;
		String val;
		
		int valint = 0;

		try {
			logger.debug("-------------------------------READING THE SPREADSHEET-------------------------------------");
			fis = new FileInputStream(fileName);
			XSSFWorkbook workbookRead = new XSSFWorkbook(fis);

			for (int k = 0; k < workbookRead.getNumberOfSheets(); k++) {
				String fileSheetName = tablename + workbookRead.getSheetName(k);
				XSSFSheet spreadsheetRead = workbookRead.getSheetAt(k);
				Iterator<Row> rowIterator = spreadsheetRead.iterator();
				StringBuffer c = new StringBuffer();
				while (rowIterator.hasNext()) {
					row = (XSSFRow) rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						// cell.setCellType(CellType.STRING);

						switch (cell.getCellType()) {

						case STRING:// System.out.print(cell.getStringCellValue()+"\t");
							val = cell.getStringCellValue();
							list.add(val);
							// System.out.println(val);
							break;

						case NUMERIC:// System.out.print(cell.getNumericCellValue()+"\t");

							if (DateUtil.isCellDateFormatted(cell)) {
								SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
								String s = (dateFormat.format(cell.getDateCellValue()));
								list.add(s);
								break;
							} else {
								valint = (int) cell.getNumericCellValue();
								list.add(valint);
								break;
							}

							// case BOOLEAN:System.out.println(cell.getBooleanCellValue());break;
						}
					}

					// System.out.println(list);
					// Create table
					if (flag == true) {
						flag = false;
						StringBuffer b = new StringBuffer();

						b.append("CREATE TABLE   IF NOT EXISTS " + fileSheetName + "( " + fileSheetName
								+ "_Id SERIAL  PRIMARY KEY,");
						excelColumnList.add(fileSheetName + "_Id");
						int len = list.size();

						for (int i = 0; i < list.size(); i++) {
							logger.debug(list.get(i));
							excelColumnList.add((String) list.get(i));
							String colName = list.get(i).toString();
							colName = colName.toLowerCase();
							if (colName.contains("date") && len == i + 1) {
								b.append(list.get(i) + " date); ");
								c.append(list.get(i) + ")");
							}

							else if (colName.contains("date")) {
								b.append(list.get(i) + " date, ");
								c.append(list.get(i) + ",");
							} else if (i == 0) {
								b.append(list.get(i) + " varchar(20)unique, ");
								c.append(list.get(i) + ",");

							} else if (len == i + 1) {
								b.append(list.get(i) + " varchar(20));");
								c.append(list.get(i) + ")");

							} else {
								b.append(list.get(i) + " varchar(20),");
								c.append(list.get(i) + ",");
							}

						}
						logger.debug(b);
						logger.debug(c);
						// System.out.println(excelColumnList);
						createTable(b);
						getColumnList(fileSheetName);
						list.clear();
						// excelColumnList.clear();

					}
					// Insert Record
					else {
						// System.out.println(excelColumnList);

						// StringBuffer s = new StringBuffer();
						StringBuffer b = new StringBuffer();
						StringBuffer updatedata = new StringBuffer();

						b.append("insert into " + fileSheetName + "(");
						b.append(c);
						b.append(" values(");
						// System.out.println(excelColumnList);
						int len = list.size();
						// System.out.println(b);
						// System.out.println(excelColumnList);
						for (int i = 0; i < list.size(); i++) {
							if (len == i + 1) {
								b.append("\'" + list.get(i) + "\')");

							} else {
								b.append("\'" + list.get(i) + "\' ,");

							}
							try {
								if (excelColumnList.size() > i + 2) {
									if (!(i == 0) && len == i + 2) {

										updatedata.append(excelColumnList.get(i + 2) + "=\'" + list.get(i + 1) + "\';");
										// System.out.println(updatedata);
									} else {
										updatedata.append(excelColumnList.get(i + 2) + "=\'" + list.get(i + 1) + "\',");
										// System.out.println(updatedata);
									}
								}
							} catch (Exception e) {
								// TODO: handle exception
								logger.error("please Fill all cell");
							}
						}

						b.append("ON CONFLICT (" + excelColumnList.get(1) + ") DO  UPDATE SET ");
						b.append(updatedata);
						logger.debug(b);

						insertTable(b);
						list.clear();

					}

					
				}
				//c=null;
				excelColumnList.clear();
				flag = true;
			}
			fis.close();

		} catch (IOException e) {

			logger.warn("Please save & close file");
		}
	}

	private void getColumnList(String fileSheetName) {
		Connection connection = null;
		try {

			Class.forName("org.postgresql.Driver");
			// System.out.println("connected");
			connection = DriverManager.getConnection(url, user, password);
			logger.debug("Successfully Connected to the database!");

		} catch (ClassNotFoundException e) {

			logger.error("Could not find the database driver ");
		} catch (SQLException e) {

			logger.error("Could not connect to the database ");
		}

		try {

			// Create a result set

			Statement statement = connection.createStatement();

			ResultSet results = statement.executeQuery("SELECT * FROM " + fileSheetName);

			// Get resultset metadata

			ResultSetMetaData metadata = results.getMetaData();

			int columnCount = metadata.getColumnCount();

			// System.out.println("test_table columns : ");

			// Get the column names; column indices start from 1
			tableColumnList.clear();
			for (int i = 1; i <= columnCount; i++) {

				String columnName = metadata.getColumnName(i);
				tableColumnList.add(metadata.getColumnName(i));
				logger.debug(columnName);

			}
			connection.close();
			// System.out.println(tableColumnList);
			excelColumnList.replaceAll(String::toLowerCase);
			tableColumnList.replaceAll(String::toLowerCase);
			List<String> uncommon = new ArrayList<>();
			for (String s : tableColumnList) {
				if (!excelColumnList.contains(s))
					uncommon.add(s);
			}
			for (String s : excelColumnList) {
				if (!tableColumnList.contains(s))
					uncommon.add(s);
			}
			// System.out.println(uncommon);

			if (uncommon.isEmpty()) {
				// System.out.println("not");

			} else {
				if (excelColumnList.size() == tableColumnList.size()) {
				} else if (excelColumnList.size() > tableColumnList.size()) {
					for (int i = 0; i < uncommon.size(); i++) {
						try {

							Class.forName("org.postgresql.Driver");
							// System.out.println("connected");
							Connection conn = DriverManager.getConnection(url, user, password);
							Statement stmt = conn.createStatement();
							String query = "ALTER TABLE " + fileSheetName + " ADD COLUMN " + uncommon.get(i)
									+ " varchar(50)";

							stmt.executeUpdate(query);
							logger.info(uncommon.get(i) + " column added.");

							// closing connection
							conn.close();

						} catch (ClassNotFoundException e) {
							// TODO Auto-generated catch block
							logger.error("Could not find the database driver ");
						} catch (SQLException e) {
							logger.error("Could not connect to the database ");
						}
					}
				} else {
					for (int i = 0; i < uncommon.size(); i++) {
						try {

							Class.forName("org.postgresql.Driver");
							// System.out.println("connected");
							Connection conn = DriverManager.getConnection(url, user, password);
							Statement stmt = conn.createStatement();
							String query = "ALTER TABLE " + fileSheetName + " DROP COLUMN " + uncommon.get(i);

							stmt.executeUpdate(query);
							logger.info(uncommon.get(i) + " column deleted.");

							// closing connection
							conn.close();

						} catch (ClassNotFoundException e) {
							// TODO Auto-generated catch block
							logger.error("Could not find the database driver ");
						} catch (SQLException e) {
							logger.error("Could not connect to the database ");
						}
					}

				}
			}

		} catch (SQLException e) {

			logger.error("Could not retrieve database metadata " + e.getMessage());
		}

	}

	private void insertTable(StringBuffer b) {
		// TODO Auto-generated method stub
		String s = b.toString();

		try {
			int result = 0;

			Class.forName("org.postgresql.Driver");
			// System.out.println("connected");
			Connection conn = DriverManager.getConnection(url, user, password);
			Statement stmt = conn.createStatement();

			result = stmt.executeUpdate(s);
			if (result == 1) {
				logger.debug("Record Inserted ");
				recordInserted++;
			} else {
				// System.out.println("Table already exists");
				// recordNotInserted++;
				// list.add(s);
			}

			conn.close();

		} catch (SQLException e) {
			recordNotInserted++;
			notInsertedRecorLlist.add(s);
			logger.error("record not inserted something is wrong");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			logger.error("Class Not Found");
		}
	}

	private void createTable(StringBuffer b) {
		// TODO Auto-generated method stub
		String s = b.toString();
		try {
			int result = 0;

			Class.forName("org.postgresql.Driver");
			logger.debug("connected");
			Connection conn = DriverManager.getConnection(url, user, password);
			Statement stmt = conn.createStatement();

			result = stmt.executeUpdate(s);
			if (result == 1) {
				logger.debug("Table created");
			} else {
				// System.out.println("Table already exists");
			}
			conn.close();

		} catch (SQLException e) {
			logger.error("Database connection problem");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			logger.error("Class Not Found");
		}
	}

	public void getAllFilesWithCertainExtension(File folder, String filterExt) {
		MyExtFilter extFilter = new MyExtFilter(filterExt);
		if (!folder.isDirectory()) {
			logger.debug("Not a folder");
		} else {
			// list out all the file name and filter by the extension
			fileNameExcel = folder.list(extFilter);

			if (fileNameExcel.length == 0) {
				logger.debug("no files end with : " + filterExt);
				return;
			}

			for (int i = 0; i < fileNameExcel.length; i++) {
				logger.debug("File :" + fileNameExcel[i]);
			}
		}
	}

	public class MyExtFilter implements FilenameFilter {

		private String ext;

		public MyExtFilter(String ext) {
			this.ext = ext;
		}

		public boolean accept(File dir, String name) {
			return (name.endsWith(ext));
		}
	}

}
