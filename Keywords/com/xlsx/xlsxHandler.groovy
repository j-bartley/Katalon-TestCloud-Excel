package com.xlsx
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

public class xlsxHandler {
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
		Path filePath = downloadsDirectory.resolve(fileName);
		String filePathString = filePath.toString();
		ExcelFactory sampleExcelFactory = new ExcelFactory();
		TestData sampleFile = sampleExcelFactory.getExcelDataWithDefaultSheet(filePathString, "Sheet1", true)

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
}
