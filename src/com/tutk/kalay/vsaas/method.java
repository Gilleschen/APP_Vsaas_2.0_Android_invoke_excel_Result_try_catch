package com.tutk.kalay.vsaas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Calendar;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import com.experitest.appium.SeeTestAndroidDriver;
import com.experitest.appium.SeeTestCapabilityType;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class method {
	int port = 8889;// Appium studio port
	int device_timeout = 60;// 60sec
	int command_timeout = 30;// 30sec
	LoadExpectResult ExpectResult = new LoadExpectResult();
	static LoadTestCase TestCase = new LoadTestCase();
	static AndroidDriver driver[] = new AndroidDriver[TestCase.DeviceInformation.deviceName.size()];
	WebDriverWait[] wait = new WebDriverWait[driver.length];
	static XSSFWorkbook workBook;
	static String appElemnt;// APP元件名稱
	static String appInput;// 輸入值
	// static int InputSec;
	String element[] = new String[driver.length];
	static int CurrentCaseNumber = 0;// 目前執行到第幾個測試案列
	XSSFSheet Sheet;

	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {

		invokeFunction();
		System.out.println("測試結束!!!!!!!!");

	}

	public static void invokeFunction() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {
		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (int i = 0; i < TestCase.StepList.size(); i++) {

			switch (TestCase.StepList.get(i).toString()) {

			case "LaunchAPP":
				methodName = "LaunchAPP";
				break;

			case "Byid_SendKey":
				methodName = "Byid_SendKey";
				appElemnt = TestCase.StepList.get(i + 1);
				appInput = TestCase.StepList.get(i + 2);
				i = i + 2;
				break;

			case "Byid_Click":
				methodName = "Byid_Click";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ByXpath_SendKey":
				methodName = "ByXpath_SendKey";
				appElemnt = TestCase.StepList.get(i + 1);
				appInput = TestCase.StepList.get(i + 2);
				i = i + 2;
				break;

			case "ByXpath_Click":
				methodName = "ByXpath_Click";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "Byid_Wait":
				methodName = "Byid_Wait";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ByXpath_Wait":
				methodName = "ByXpath_Wait";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "HideKeyboard":
				methodName = "HideKeyboard";
				break;

			case "Byid_Result":
				methodName = "Byid_Result";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ByXpath_Result":
				methodName = "ByXpath_Result";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "Sleep":
				methodName = "Sleep";
				appInput = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ScreenShot":
				methodName = "ScreenShot";
				break;
			case "ResetAPP":
				methodName = "ResetAPP";
				break;

			case "QuitAPP":
				methodName = "QuitAPP";
				break;

			}

			Method method;
			method = c.getMethod(methodName, new Class[] {});
			method.invoke(c.newInstance());

		}
	}

	public void Byid_Result() {
		boolean result[] = new boolean[driver.length];// 未給定Boolean值，預設為False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {
			try {
				element[i] = driver[i].findElement(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt))
						.getText();
			} catch (Exception e) {
				element[i] = "ERROR";// 找不到該物件，回傳Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
				ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
				for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
					if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
						result[i] = true;
						break;
					} else {
						result[i] = false;
					}
				}
			}
		}
		// 開啟Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TestReport\\TestReport.xlsm"));

		} catch (Exception e) {
			;
		}
		for (int i = 0; i < driver.length; i++) {
			Sheet = workBook.getSheet(TestCase.DeviceInformation.deviceName.get(i) + "_TestReport");// 指定某台裝置的TestReport
																									// sheet
			if (ErrorResult[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error!!");
			} else if (result[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
			} else if (result[i] == false) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
			}
		}
		// 執行寫入Excel後的存檔動作
		try {
			FileOutputStream out = new FileOutputStream(new File("C:\\TestReport\\TestReport.xlsm"));

			workBook.write(out);
			out.close();
			workBook.close();

		} catch (Exception e) {
			;
		}

		CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void ByXpath_Result() {
		boolean result[] = new boolean[driver.length];// 未給定Boolean值，預設為False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {

			try {
				element[i] = driver[i].findElement(By.xpath(appElemnt)).getAttribute("content-desc");
			} catch (Exception e) {
				element[i] = "ERROR";// 找不到該物件，回傳Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
				ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
				for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
					if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
						result[i] = true;
						break;
					} else {
						result[i] = false;
					}
				}
			}

		}
		// 開啟Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TestReport\\TestReport.xlsm"));

		} catch (Exception e) {
			;
		}
		for (int i = 0; i < driver.length; i++) {
			Sheet = workBook.getSheet(TestCase.DeviceInformation.deviceName.get(i) + "_TestReport");// 指定某台裝置的TestReport
																									// sheet
			if (ErrorResult[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error!!");
			} else if (result[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
			} else if (result[i] == false) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
			}
		}
		// 執行寫入Excel後的存檔動作
		try {
			FileOutputStream out = new FileOutputStream(new File("C:\\TestReport\\TestReport.xlsm"));

			workBook.write(out);
			out.close();
			workBook.close();

		} catch (Exception e) {
			;
		}

		CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void Byid_Wait() {
		// WebDriverWait[] wait = new WebDriverWait[driver.length];
		for (int i = 0; i < driver.length; i++) {
			wait[i] = new WebDriverWait(driver[i], device_timeout);
			try {
				wait[i].until(ExpectedConditions
						.presenceOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)));
			} catch (Exception e) {
				System.out.println("Error");
			}
		}
	}

	public void ByXpath_Wait() {
		// WebDriverWait[] wait = new WebDriverWait[driver.length];
		for (int i = 0; i < driver.length; i++) {
			wait[i] = new WebDriverWait(driver[i], device_timeout);
			try {
				wait[i].until(ExpectedConditions.presenceOfElementLocated(By.xpath(appElemnt)));
			} catch (Exception e) {
				;
			}
		}
	}

	public void Byid_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				driver[i].findElement(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt));

			} catch (Exception e) {
				;
			}
		}
	}

	public void Byid_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				driver[i].findElement(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)).click();
			} catch (Exception e) {
				;
			}

		}
	}

	public void ByXpath_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				driver[i].findElement(By.xpath(appElemnt)).sendKeys(appInput);
			} catch (Exception e) {
				;
			}

		}
	}

	public void ByXpath_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				driver[i].findElement(By.xpath(appElemnt)).click();
			} catch (Exception e) {
				;
			}

		}
	}

	public void HideKeyboard() {
		for (int i = 0; i < driver.length; i++) {

			try {
				driver[i].hideKeyboard();
			} catch (Exception ex) {
				;
			}
		}
	}

	public void Sleep() {
		String NewString = "";// 新字串
		char[] r = { '.' };// 小數點字元
		char[] c = appInput.toCharArray();// 將字串轉成字元陣列
		for (int i = 0; i < c.length; i++) {
			if (c[i] != r[0]) {// 判斷字元是否為小數點
				NewString = NewString + c[i];// 否，將字元組合成新字串
			} else {
				break;// 是，跳出迴圈
			}
		}
		// int t = Integer.valueOf(s) * 1000;
		try {
			System.out.println("[driver] [start] sleep(): " + NewString + " second...");
			Thread.sleep(Integer.valueOf(NewString) * 1000);// 將字串轉成整數
			System.out.println("[driver] [end] sleep");
		} catch (Exception e) {
			;
		}
	}

	public void ScreenShot() {
		Calendar date = Calendar.getInstance();
		String month = Integer.toString(date.get(Calendar.MONTH) + 1);
		String day = Integer.toString(date.get(Calendar.DAY_OF_MONTH));
		String hour = Integer.toString(date.get(Calendar.HOUR_OF_DAY));
		String min = Integer.toString(date.get(Calendar.MINUTE));
		String sec = Integer.toString(date.get(Calendar.SECOND));
		for (int i = 0; i < driver.length; i++) {
			File screenShotFile = (File) driver[i].getScreenshotAs(OutputType.FILE);

			try {
				FileUtils.copyFile(screenShotFile, new File(
						TestCase.CaseList.get(CurrentCaseNumber) + "_" + month + day + hour + min + sec + ".jpg"));
				System.out.println("[Log] " + "ScreenShoot Successfully!! (CaseName+Month+Day+Hour+Minus+Second)");
			} catch (IOException e) {
				;
			}
		}
		CurrentCaseNumber = CurrentCaseNumber + 1;
	}

	public void QuitAPP() {
		for (int i = 0; i < driver.length; i++) {
			driver[i].quit();
		}
	}

	public void ResetAPP() {
		for (int i = 0; i < driver.length; i++) {
			driver[i].resetApp();
		}
	}

	public void LaunchAPP() {
		DesiredCapabilities cap[] = new DesiredCapabilities[TestCase.DeviceInformation.deviceName.size()];

		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			cap[i] = new DesiredCapabilities();
		}

		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {

			for (int j = i; j < TestCase.DeviceInformation.platformVersion.size(); j++) {

				cap[i].setCapability(SeeTestCapabilityType.WAIT_FOR_DEVICE_TIMEOUT_MILLIS, device_timeout * 1000);
				cap[i].setCapability(MobileCapabilityType.UDID, TestCase.DeviceInformation.deviceName.get(i));
				cap[i].setCapability(AndroidMobileCapabilityType.APP_PACKAGE, TestCase.DeviceInformation.appPackage);
				cap[i].setCapability(AndroidMobileCapabilityType.APP_ACTIVITY, TestCase.DeviceInformation.appActivity);
				cap[i].setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, command_timeout);

				try {
					driver[j] = new SeeTestAndroidDriver<>(new URL("http://localhost:" + port + "/wd/hub"), cap[j]);
				} catch (MalformedURLException e) {
					;
				}
				break;
			}
			// port = port + 2;
		}
	}

}
