package com.katalon.testcloud
import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.remote.RemoteWebDriver
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import internal.GlobalVariable
import java.io.File;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import com.kms.katalon.core.testdata.ExcelData as ExcelData
import com.kms.katalon.core.testdata.reader.ExcelFactory as ExcelFactory
import com.kms.katalon.core.testdata.AbstractTestData as AbstractTestData
import java.util.regex.Matcher
import java.util.regex.Pattern

public class FileExecutor {
	/**
	 * Download file content using base64 encoding.
	 *
	 * @param the given file name
	 * @return the file content in base64
	 */
	@Keyword
	def getFileContent(String fileName) {
		try {
			Path userHome = Paths.get(System.getProperty("user.home"));
			Path downloadsDirectory = userHome.resolve("Downloads");
			Path filePath = downloadsDirectory.resolve(fileName);
			File file = filePath.toFile();
			byte[] fileBytes = Files.readAllBytes(file.toPath());
			return Base64.getEncoder().encodeToString(fileBytes);
		} catch (Exception e) {
			throw new Exception('Failed to execute TestCloud Keyword: FileExecutor.getFileContent - Error Code: TCKW301');
		}
	}
	/**
	 * Download .xlsx file content
	 *
	 * @param the given file name
	 * @return the file content in base64
	 */
	@Keyword
	def getExcelFileContent(String fileName) {
		Path userHome = Paths.get(System.getProperty("user.home"));
		Path downloadsDirectory = userHome.resolve("Downloads");
		File dir = new File(downloadsDirectory.toString())
		File[] dirContents = dir.listFiles()
		int counter = 0
		int fileCount = dirContents.size()
		String filePath = ""
		while (counter < fileCount) {
			Pattern p1 = Pattern.compile('(.*)Content_Analysis-(.*)')
			Matcher m1 = p1.matcher(dirContents[counter].toString())
			if (m1.find()) {
				filePath = dirContents[counter]
				break
			}
			counter = counter + 1
		}
		ExcelFactory sampleExcelFactory = new ExcelFactory();
		TestData sampleFile = sampleExcelFactory.getExcelDataWithDefaultSheet(filePath, "Sheet1", true)

		// At this point we now have the .xlsx file saved as TestData, and can manipulate it how we chose
		// For example, we could now pass the sampleFile TestData to another keyword for parsing or another function
		// In this example, instead of calling another keyword, we are iterating through the entire file and printing the values of the cells
		int columnCount = sampleFile.getColumnNumbers();

		int rowCount = sampleFile.getRowNumbers();

		int rowIterate = 1;

		int columnIterate = 1;

		String cellValue = '';

		while (rowIterate <= rowCount) {
			while (columnIterate <= columnCount) {
				cellValue = sampleFile.getValue(columnIterate, rowIterate);

				println("Cell Value is " + cellValue + " at Column " + columnIterate + " and at Row " + rowIterate);
				columnIterate = (columnIterate + 1);
			}

			columnIterate = 1;

			rowIterate = (rowIterate + 1);
		}
	}
	/**
	 * Retrieve file metadata.
	 *
	 * @param the given file name
	 * @return file metadata
	 */
	@Keyword
	def getFileDescriptor(String fileName) {
		try {
			Path userHome = Paths.get(System.getProperty("user.home"));
			Path downloadsDirectory = userHome.resolve("Downloads");
			Path filePath = downloadsDirectory.resolve(fileName);
			File file = filePath.toFile();
			long fileSize = file.length();
			Map< String, Object > fileDescriptor = new HashMap<>();
			long lastModifiedTimestamp = file.lastModified();
			FileTime fileTime = FileTime.fromMillis(lastModifiedTimestamp);
			Instant instant = fileTime.toInstant();
			ZonedDateTime utcDateTime = instant.atZone(ZoneId.of("UTC"));
			DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss'Z'");
			String utcDateString = utcDateTime.format(formatter);
			fileDescriptor.put("modified_time", utcDateString);
			fileDescriptor.put("name", file.getName());
			fileDescriptor.put("size", fileSize);
			return fileDescriptor;
		} catch (Exception e) {
			throw new Exception('Failed to execute TestCloud Keyword: FileExecutor.getFileDescriptor - Error Code: TCKW302');
		}
	}
	/**
	 * Check if file name exists.
	 *
	 * @param the given file name
	 * @return true if the file exists
	 */
	@Keyword
	def exist(String fileName) {
		try {
			Path userHome = Paths.get(System.getProperty("user.home"));
			Path downloadsDirectory = userHome.resolve("Downloads");
			Path filePath = downloadsDirectory.resolve(fileName);
			File file = filePath.toFile();
			return file.exists();
		} catch (Exception e) {
			throw new Exception('Failed to execute TestCloud Keyword: FileExecutor.exist - Error Code: TCKW303');
		}
	}
}