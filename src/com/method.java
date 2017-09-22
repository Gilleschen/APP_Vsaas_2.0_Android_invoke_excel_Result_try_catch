package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import java.net.URL;
import java.util.Calendar;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.ScreenOrientation;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import com.experitest.appium.SeeTestAndroidDriver;
import com.experitest.appium.SeeTestCapabilityType;

import io.appium.java_client.android.AndroidKeyCode;
import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class method {
	int port = 4725;// Appium studio port
	int device_timeout = 60;// 60sec
	int command_timeout = 30;// 30sec
	LoadExpectResult ExpectResult = new LoadExpectResult();
	static LoadTestCase TestCase = new LoadTestCase();

	static String CaseErrorList[][] = new String[TestCase.CaseList.size()][TestCase.DeviceInformation.deviceName
			.size()];// �����U�רҩ�U�˸m�����O���G (2���}�C)CaseErrorList[CaseList][Devices]
	String ErrorList[] = new String[TestCase.DeviceInformation.deviceName.size()];// �����U�˸m�����O���G

	static SeeTestAndroidDriver driver[] = new SeeTestAndroidDriver[TestCase.DeviceInformation.deviceName.size()];
	WebDriverWait[] wait = new WebDriverWait[driver.length];
	static XSSFWorkbook workBook;
	static String appElemnt;// APP����W��
	static String appInput;// ��J��
	static String appInputXpath;// ��J�Ȫ�Xpath�榡
	static String toElemnt;// APP����W��
	static int startx, starty, endx, endy;// Swipe���ʮy��
	static int iterative;// �e���ưʦ���
	static String scroll;// �e�����ʤ�V
	static String appElemntarray;// �j�M���h����������
	String element[] = new String[driver.length];
	static int CurrentCaseNumber = -1;// �ثe�����ĴX�Ӵ��ծצC
	static Boolean CommandError = true;// �P�w���檺���O�O�_�X�{���~�Fture�����T�Ffalse�����~
	XSSFSheet Sheet;

	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException, IOException {
		initial();
		invokeFunction();
		System.out.println("���յ���!!!!!!!!");
		Process proc = Runtime.getRuntime().exec("explorer C:\\TUTK_QA_TestTool\\TestReport");// �}��TestReport��Ƨ�

	}

	public static void initial() {// ��l��CaseErrorList�x�}
		for (int i = 0; i < CaseErrorList.length; i++) {
			for (int j = 0; j < CaseErrorList[i].length; j++) {
				CaseErrorList[i][j] = "";//��J�Ŧr��A�קK���ȮɡA�X�{���~
			}
		}
	}

	public static void invokeFunction() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {
		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (int CurrentCase = 0; CurrentCase < TestCase.StepList.size(); CurrentCase++) {
			CommandError = true;// �w�]CommandError��True

			for (int CurrentCaseStep = 0; CurrentCaseStep < TestCase.StepList.get(CurrentCase)
					.size(); CurrentCaseStep++) {
				if (!CommandError) {
					break;// �Y�ثe���ծרҥX�{CommandError=false�A�h���X�ثe�רҨð���U�@�Ӯר�
				}
				switch (TestCase.StepList.get(CurrentCase).get(CurrentCaseStep).toString()) {

				case "LaunchAPP":
					methodName = "LaunchAPP";
					break;

				case "Byid_SendKey":
					methodName = "Byid_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "Byid_Click":
					methodName = "Byid_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_Swipe":
					methodName = "Byid_Swipe";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					toElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "ByXpath_SendKey":
					methodName = "ByXpath_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "ByXpath_Click":
					methodName = "ByXpath_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Swipe":
					methodName = "ByXpath_Swipe";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					toElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "Byid_Wait":
					methodName = "Byid_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Wait":
					methodName = "ByXpath_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "HideKeyboard":
					methodName = "HideKeyboard";
					break;

				case "Byid_Result":
					methodName = "Byid_Result";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Result":
					methodName = "ByXpath_Result";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Sleep":
					methodName = "Sleep";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ScreenShot":
					methodName = "ScreenShot";
					break;

				case "Orientation":
					methodName = "Orientation";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Swipe":
					methodName = "Swipe";
					startx = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1));
					starty = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2));
					endx = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					endy = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 4));
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 5));
					CurrentCaseStep = CurrentCaseStep + 5;
					break;

				case "ByXpath_Swipe_Vertical":
					methodName = "ByXpath_Swipe_Vertical";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					CurrentCaseStep = CurrentCaseStep + 3;
					break;

				case "ByXpath_Swipe_Horizontal":
					methodName = "ByXpath_Swipe_Horizontal";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					CurrentCaseStep = CurrentCaseStep + 3;
					break;

				case "ByXpath_Swipe_FindText_Click_Android":
					methodName = "ByXpath_Swipe_FindText_Click_Android";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					appElemntarray = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 4);
					appInputXpath = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 5);
					CurrentCaseStep = CurrentCaseStep + 5;
					break;

				case "Back":
					methodName = "Back";
					break;

				case "Home":
					methodName = "Home";
					break;

				case "Power":
					methodName = "Power";
					break;

				case "Menu":
					methodName = "Menu";
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
	}

	public void Byid_Result() {
		boolean result[] = new boolean[driver.length];// �����wBoolean�ȡA�w�]��False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				element[i] = wait[i].until(ExpectedConditions
						.visibilityOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)))
						.getText();
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				element[i] = "ERROR";// �䤣��Ӫ���A�^��Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// �^�Ǵ��ծרҲM�檺�W�ٵ�ExpectResult.LoadExpectResult�A�æs����浲�G��ResultList�M��
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
		SubMethod_Result(ErrorResult, result);// �I�ssubmethod_result�x�s���յ��G��Excel
		// CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void ByXpath_Result() {
		boolean result[] = new boolean[driver.length];// �����wBoolean�ȡA�w�]��False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {

			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)))
						.getAttribute("content-desc");
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				element[i] = "ERROR";// �䤣��Ӫ���A�^��Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// �^�Ǵ��ծרҲM�檺�W�ٵ�ExpectResult.LoadExpectResult�A�æs����浲�G��ResultList�M��
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
		SubMethod_Result(ErrorResult, result);

		// CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void Byid_Wait() {

		for (int i = 0; i < driver.length; i++) {
			wait[i] = new WebDriverWait(driver[i], device_timeout);
			try {
				wait[i].until(ExpectedConditions
						.presenceOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)));
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}
	}

	public void ByXpath_Wait() {

		for (int i = 0; i < driver.length; i++) {
			wait[i] = new WebDriverWait(driver[i], device_timeout);
			try {
				wait[i].until(ExpectedConditions.presenceOfElementLocated(By.xpath(appElemnt)));
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}
	}

	public void Byid_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions
						.visibilityOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)))
						.sendKeys(appInput);
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;

			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}
	}

	public void Byid_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions
						.visibilityOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)))
						.click();
				ErrorList[i] = "Pass";// �x�s��i�x�]��command���G�A���\����Click�h�A�s�JPass||�|��
										// ���N1�GErrorList=[Pass];���N2�GErrorList=[Pass,Pass]
				CaseErrorList[CurrentCaseNumber] = ErrorList;// �x�s��i�x�]�ư����CurrentCaseNumber�ӮרҤ�command���G||�|��
																// ���N1�GCaseErrorList=[[Pass]];���N2�GCaseErrorList=[[Pass,Pass]]
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false

			}
		}

	}

	public void ByXpath_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).sendKeys(appInput);
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}
	}

	public void ByXpath_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).click();
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
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
		String NewString = "";// �s�r��
		char[] r = { '.' };// �p���I�r��
		char[] c = appInput.toCharArray();// �N�r���ন�r���}�C
		for (int i = 0; i < c.length; i++) {
			if (c[i] != r[0]) {// �P�_�r���O�_���p���I
				NewString = NewString + c[i];// �_�A�N�r���զX���s�r��
			} else {
				break;// �O�A���X�j��
			}
		}

		try {
			System.out.println("[driver] [start] Sleep(): " + NewString + " second...");
			Thread.sleep(Integer.valueOf(NewString) * 1000);// �N�r���ন���
			System.out.println("[driver] [end] Sleep");
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
				FileUtils.copyFile(screenShotFile, new File("C:\\TUTK_QA_TestTool\\TestReport\\"
						+ TestCase.CaseList.get(CurrentCaseNumber) + "_" + month + day + hour + min + sec + ".jpg"));
				System.out.println("[Log] " + "ScreenShot Successfully!! (CaseName+Month+Day+Hour+Minus+Second)");
			} catch (IOException e) {
				System.out.println("[Error]Fail to ScreenShot");
			}
		}
	}

	public void Orientation() {

		for (int i = 0; i < driver.length; i++) {
			if (appInput.equals("Landscape")) {
				driver[i].rotate(ScreenOrientation.LANDSCAPE);
			} else if (appInput.equals("Portrait")) {
				driver[i].rotate(ScreenOrientation.PORTRAIT);
			}
		}
	}

	public void QuitAPP() {
		for (int j = 0; j < driver.length; j++) {
			driver[j].quit();// ���}APP��A�g�J���յ��GPass��Error

			// �}��Excel
			try {
				workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
			} catch (Exception e) {
				System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
			}
			for (int i = 0; i < driver.length; i++) {

				if (TestCase.DeviceInformation.deviceName.get(i).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
					char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
					TestCase.DeviceInformation.deviceName.get(i).toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
					Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// �ھ�NewUdid�A���w�Y�x�˸m��TestReport
																						// sheet
				} else {
					Sheet = workBook.getSheet(TestCase.DeviceInformation.deviceName.get(i).toString() + "_TestReport");// ���w�Y�x�˸m��TestReport
																														// sheet
				}

				if (CaseErrorList[CurrentCaseNumber][i].equals("Pass")) {// ���XCaseErrorList����CurrentCaseNumber�Ӵ���������i�x��ʸ˸m�����G
					Sheet.getRow(CurrentCaseNumber + 1).getCell(1).setCellValue("Pass");// ��J��i�x��ʸ˸m����CurrentCaseNumber�Ӵ������GPass
				}

			}
			// ����g�JExcel�᪺�s�ɰʧ@
			try {
				FileOutputStream out = new FileOutputStream(
						new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
				workBook.write(out);
				out.close();
				workBook.close();
			} catch (Exception e) {
				System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
			}

		}
	}

	public void ResetAPP() {
		for (int i = 0; i < driver.length; i++) {
			driver[i].resetApp();
		}
	}

	public void LaunchAPP() {
		DesiredCapabilities cap[] = new DesiredCapabilities[TestCase.DeviceInformation.deviceName.size()];
		CurrentCaseNumber = CurrentCaseNumber + 1;
		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			cap[i] = new DesiredCapabilities();
		}

		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {

			for (int j = i; j < TestCase.DeviceInformation.platformVersion.size(); j++) {

				cap[i].setCapability(SeeTestCapabilityType.WAIT_FOR_DEVICE_TIMEOUT_MILLIS, device_timeout * 1000);
				cap[i].setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, command_timeout);
				cap[i].setCapability(MobileCapabilityType.UDID, TestCase.DeviceInformation.deviceName.get(i));
				cap[i].setCapability(AndroidMobileCapabilityType.APP_PACKAGE, TestCase.DeviceInformation.appPackage);
				cap[i].setCapability(AndroidMobileCapabilityType.APP_ACTIVITY, TestCase.DeviceInformation.appActivity);

				cap[i].setCapability(SeeTestCapabilityType.REPORT_FORMAT, "xml");
				cap[i].setCapability(SeeTestCapabilityType.REPORT_DIRECTORY, "C:\\TUTK_QA_TestTool\\TestReport");// Report���|
				cap[i].setCapability(SeeTestCapabilityType.TEST_NAME, TestCase.CaseList.get(CurrentCaseNumber));// TestCase�W��

				try {
					driver[j] = new SeeTestAndroidDriver(new URL("http://localhost:" + port + "/wd/hub"), cap[j]);

				} catch (Exception e) {
					System.out.print("[Error] Can't find UDID: " + TestCase.DeviceInformation.deviceName.get(i));
					System.out.println(" or can not find appPackage: " + TestCase.DeviceInformation.appPackage);
				}
				break;
			}
			// port = port + 2;
		}
	}

	public void Back() {
		for (int i = 0; i < driver.length; i++) {
			driver[i].pressKeyCode(AndroidKeyCode.BACK);
		}
	}

	public void Home() {
		for (int i = 0; i < driver.length; i++) {

			driver[i].pressKeyCode(AndroidKeyCode.HOME);
		}
	}

	public void Power() {
		for (int i = 0; i < driver.length; i++) {

			driver[i].pressKeyCode(AndroidKeyCode.KEYCODE_POWER);
		}
	}

	public void Menu() {
		for (int i = 0; i < driver.length; i++) {
			driver[i].pressKeyCode(AndroidKeyCode.MENU);
			// driver[i].deviceAction("Recent Apps");
		}
	}

	public void ByXpath_Swipe() {
		Point p1, p2;// p1 ���_�I;p2�����I

		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				p2 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(toElemnt))).getLocation();
				p1 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).getLocation();
				driver[i].swipe(p1.x, p1.y, p1.x, p1.y - (p1.y - p2.y), 1000);
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.print("[Error] Can't find " + appElemnt);
				System.out.println(" or Can't find " + toElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}
	}

	public void Byid_Swipe() {
		Point p1, p2;// p1 ���_�I;p2�����I

		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				p2 = wait[i].until(ExpectedConditions
						.visibilityOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + toElemnt)))
						.getLocation();
				p1 = wait[i].until(ExpectedConditions
						.visibilityOfElementLocated(By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)))
						.getLocation();

				driver[i].swipe(p1.x, p1.y, p1.x, p1.y - (p1.y - p2.y), 1000);
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.print("[Error] Can't find " + appElemnt);
				System.out.println(" or Can't find " + toElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}
	}

	public void Swipe() {
		for (int i = 0; i < driver.length; i++) {
			for (int j = 0; j < iterative; j++) {
				driver[i].swipe(startx, starty, endx, endy, 500);
			}
		}
	}

	// ���[�Jinvoke function
	/*
	 * public void Byid_Swipe_Vertical() { Point p;// ����y�� Dimension s;// ����j�p
	 * WebElement e; for (int i = 0; i < driver.length; i++) { try { wait[i] =
	 * new WebDriverWait(driver[i], command_timeout); e =
	 * wait[i].until(ExpectedConditions
	 * .visibilityOfElementLocated(By.id(TestCase.DeviceInformation.appPackage +
	 * ":id/" + appElemnt))); // e = //
	 * driver[i].findElement(By.id(TestCase.DeviceInformation.appPackage // +
	 * ":id/" + appElemnt)); s = e.getSize(); p = e.getLocation(); int errorX =
	 * (int) Math.round(s.width * 0.01); int errorY = (int) Math.round(s.height
	 * * 0.01); for (int j = 0; j < iterative; j++) { if (scroll.equals("DOWN"))
	 * {// �e���V�U���� driver[i].swipe(p.x + errorX, p.y + s.height - errorY, p.x +
	 * errorX, p.y + errorY, 1000); } else if (scroll.equals("UP")) {// �e���V�W����
	 * driver[i].swipe(p.x + errorX, p.y + errorY, p.x + errorX, p.y + s.height
	 * - errorY, 1000); } } } catch (Exception w) { System.out.println(
	 * "[Error] Can't find " + TestCase.DeviceInformation.appPackage + ":id/" +
	 * appElemnt); }
	 * 
	 * } }
	 */

	public void ByXpath_Swipe_Vertical() {
		Point p;// ����y��
		Dimension s;// ����j�p
		WebElement e;

		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				e = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.width * 0.01);
				int errorY = (int) Math.round(s.height * 0.01);
				for (int j = 0; j < iterative; j++) {
					if (scroll.equals("DOWN")) {// �e���V�U����
						driver[i].swipe(p.x + errorX, p.y + s.height - errorY, p.x + errorX, p.y + errorY, 1000);
					} else if (scroll.equals("UP")) {// �e���V�W����
						driver[i].swipe(p.x + errorX, p.y + errorY, p.x + errorX, p.y + s.height - errorY, 1000);
					}
				}
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception w) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}

		}
	}

	// ���[�Jinvoke function
	/*
	 * public void Byid_Swipe_Horizontal() { Point p;// ����y�� Dimension s;// ����j�p
	 * for (int i = 0; i < driver.length; i++) { try { wait[i] = new
	 * WebDriverWait(driver[i], command_timeout); s =
	 * driver[i].findElement(By.id(TestCase.DeviceInformation.appPackage +
	 * ":id/" + appElemnt)).getSize(); p =
	 * driver[i].findElement(By.id(TestCase.DeviceInformation.appPackage +
	 * ":id/" + appElemnt)) .getLocation(); int errorX = (int)
	 * Math.round(s.getWidth() * 0.01); int errorY = (int)
	 * Math.round(s.getHeight() * 0.01); for (int j = 0; j < iterative; j++) {
	 * if (scroll.equals("RIGHT")) {// �e���V�k���� (�[�ݵe�����褺�e) driver[i].swipe(p.x +
	 * errorX, p.y + errorY, p.x + s.width - errorX, p.y + errorY, 1000); } else
	 * if (scroll.equals("LEFT")) {// �e���V������ (�[�ݵe���k�褺�e) driver[i].swipe(p.x +
	 * s.width - errorX, p.y + errorY, p.x + errorX, p.y + errorY, 1000); } } }
	 * catch (Exception w) { System.out.print("[Error] Can't find " +
	 * appElemnt); }
	 * 
	 * } }
	 */
	public void ByXpath_Swipe_Horizontal() {
		Point p;// ����y��
		Dimension s;// ����j�p
		WebElement e;
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				e = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));
				// e = driver[i].findElement(By.xpath(appElemnt));
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.getWidth() * 0.01);
				int errorY = (int) Math.round(s.getHeight() * 0.01);
				for (int j = 0; j < iterative; j++) {
					if (scroll.equals("RIGHT")) {// �e���V�k���� (�[�ݵe�����褺�e)
						driver[i].swipe(p.x + errorX, p.y + errorY, p.x + s.width - errorX, p.y + errorY, 1000);
					} else if (scroll.equals("LEFT")) {// �e���V������ (�[�ݵe���k�褺�e)
						driver[i].swipe(p.x + s.width - errorX, p.y + errorY, p.x + errorX, p.y + errorY, 1000);
					}
				}
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception w) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}

		}
	}

	// ���[�JInvkoe function
	// ByXpath_Swipe_FindText_Click_Android_2���I
	// �G���M�w�w�����w����A������dispaly�X��(�Q�צ�A�ײ��ݤ���)�A�]������Click()��AAPP�欰�i��|�X���F�ѨM�覡�G�ϥ�wait.until(ExpectedConditions.visibilityOfElementLocated())
	// visibilityOfElementLocated��ܪ��줸����ܦb�ù��W��(�ײ��ݪ���)�A�A����U�@�B�ʧ@�F���I�A������Timeout�ɶ�(WebDriverWait��timeout�ɶ�+MobileCapabilityType.NEW_COMMAND_TIMEOUT��timeout�ɶ�)
	/*
	 * public void ByXpath_Swipe_FindText_Click_Android_2() {
	 * 
	 * for (int j = 0; j < driver.length; j++) { try { List<WebElement>
	 * targetlist = driver[j].findElementsByXPath(appElemntarray);//
	 * �n�j�M���h����������M�� String lastelement = targetlist.get(targetlist.size() -
	 * 1).getText().toString();// �n�j�M���h����������M�椤�A�̫�@���r��(�ت��G�P�_�O�_�Ҧ���Ƴ��j�M�L) for (int
	 * i = 0; i < targetlist.size(); i++) {
	 * 
	 * if ((targetlist.get(i).getText().toString()).equals(appInput)) {//
	 * �Ytargetelement�btargetlist�M�椤�A�h�I��targetelement WebElement target =
	 * targetlist.get(i); target.click(); break; }
	 * 
	 * if (i == targetlist.size() - 1) {// �Ytargetlist���̫�@����Ƥ�粒��A�h�i��Srcoll�즲
	 * 
	 * Point p;// ����y�� Dimension s;// ����j�p WebElement e; e =
	 * driver[j].findElement(By.xpath(appElemnt));// Scroll // bar //
	 * �w�����g�bFor(j)�j��~�A�i�`�ٮɶ� s = e.getSize(); p = e.getLocation(); int errorX =
	 * (int) Math.round(s.width * 0.1); int errorY = (int) Math.round(s.height *
	 * 0.1);
	 * 
	 * if (scroll.equals("DOWN")) { driver[j].swipe(p.x + errorX, s.height -
	 * errorY, p.x + errorX, p.y + errorY, 1000);// �V�U���� } else if
	 * (scroll.equals("UP")) { driver[j].swipe(p.x + errorX, p.y + errorY, p.x +
	 * errorX, s.height - errorY, 1000);// �V�W���� } else if
	 * (scroll.equals("LEFT")) { driver[j].swipe(s.width - errorX, p.y + errorY,
	 * p.x + errorX, p.y + errorY, 1000);// �e���V������(�[�ݵe���k�褺�e) } else if
	 * (scroll.equals("RIGHT")) { driver[j].swipe(p.x + errorX, p.y + errorY,
	 * s.width - errorX, p.y + errorY, 1000);// �e���V�k����(�[�ݵe�����褺�e) }
	 * 
	 * targetlist.clear();// �M��targetlist targetlist =
	 * driver[j].findElementsByXPath(appElemntarray);// �A�����o�stargetlist
	 * 
	 * if (lastelement.equals(targetlist.get(targetlist.size() -
	 * 1).getText().toString())) {// �P�_�stargetlist�̫�@����ƬO�_�Plastelement�ۦP
	 * System.out.println("�䤣��" + appInput);// �Y�ۦP�A���Srcoll�w�즲�̫ܳ�A�h�L�X�䤣��A���Xfor
	 * break; } else {// �Y���ۦP�A�h��slastelement�̫ᵧ��� lastelement =
	 * targetlist.get(targetlist.size() - 1).getText().toString(); } i = -1;//
	 * �Oi=-1(�ت��G�A������for) } }
	 * 
	 * } catch (Exception w) { System.out.print("[Error] Can't find " +
	 * appElemntarray); } } }
	 */
	// ByXpath_Swipe_FindText_Click_Android���I�G1.�j�M���r�ꤣ�i���� 2.�j�M5�����S��줸��A�h����j�M
	public void ByXpath_Swipe_FindText_Click_Android() {

		for (int j = 0; j < driver.length; j++) {
			int SearchNumber = 0;// �j�M����
			Point ScrollBarP;// ���b����y��
			Dimension ScrollBarS;// ���b���󤧪��μe
			WebElement ScrollBar;// ���b����

			try {
				wait[j] = new WebDriverWait(driver[j], command_timeout);
				ScrollBar = wait[j].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));// ���b����
				ScrollBarS = ScrollBar.getSize();// ���b���󪺪��μe
				ScrollBarP = ScrollBar.getLocation();// ���b���y��
				int errorX = (int) Math.round(ScrollBarS.width * 0.1);
				int errorY = (int) Math.round(ScrollBarS.height * 0.1);
				List<WebElement> targetlist = wait[j]
						.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(appElemntarray)));// �n�j�M���h����������M��

				for (int i = 0; i < targetlist.size(); i++) {

					if ((targetlist.get(i).getText().toString()).equals(appInput)) {// �Ytargetelement�btargetlist�M�椤�A�h�I��targetelement
						WebElement targetElement;// �ǳƷj�M������
						Point targetElementP;// �ǳƷj�M�����󤧮y��
						Dimension targetElementS;// �ǳƷj�M�����󤧪��μe

						targetElement = wait[j]
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appInputXpath)));

						targetElementP = targetElement.getLocation();// �ǳƷj�M���󪺮y��
						targetElementS = targetElement.getSize();// �ǳƷj�M���󪺪��μe

						switch (scroll.toString()) {

						case "DOWN":
							if (targetElementP.y > ScrollBarP.y + ScrollBarS.height) {// �Y�j�M����y�y�Фj����b�d��A��ܷj�M�������UI�Q���b�B��
								driver[j].swipe(targetElementP.x, ScrollBarS.height + ScrollBarP.y - errorY,
										targetElementP.x, ScrollBarP.y + errorY, 1000);
							} else if (targetElementP.y + targetElementS.height == ScrollBarP.y + ScrollBarS.height) {// �Y�j�M����y�y�лP�e���`�M������b���סA��ܷj�M���󪺳���UI�Q���b�B��
								driver[j].swipe(targetElementP.x - errorY, targetElementP.y, targetElementP.x,
										ScrollBarP.y + errorY, 1000);
							}
							break;

						case "UP":
							if (targetElementP.y + targetElementS.height < ScrollBarP.y) {// �Y�j�M���󪺳̤jy�y�Фp����by�y�СA��ܷj�M�������UI�Q���b�B��
								driver[j].swipe(targetElementP.x, ScrollBarP.y + errorY, targetElementP.x,
										ScrollBarS.height + ScrollBarP.y - errorY, 1000);
							} else {// �Ϥ��A�Y�j�M���󪺳̤jy�y�Фj����by�y�СA��ܷj�M�������UI�Q���b�B��
								driver[j].swipe(targetElementP.x, ScrollBarP.y + errorY, targetElementP.x,
										ScrollBarP.y + ScrollBarS.height - errorY, 1000);
							}
							break;

						case "LEFT":// �e���V������(�[�ݵe���k�褺�e)
							if (targetElementP.x > ScrollBarP.x + ScrollBarS.width) {// �Y�j�M����x�y�Фj����b�d��A��ܷj�M�������UI�Q���b�B��
								driver[j].swipe(ScrollBarP.x + ScrollBarS.width - errorX, targetElementP.y,
										ScrollBarP.x + errorX, targetElementP.y, 1000);
							} else if (targetElementP.x + targetElementS.width == ScrollBarP.x + ScrollBarS.width) {// �Y�j�M����x�y�лP�e���`�M������b�e�סA��ܷj�M���󪺳���UI�Q���b�B��
								driver[j].swipe(targetElementP.x - errorX, targetElementP.y, ScrollBarP.x + errorX,
										targetElementP.y, 1000);
							}
							break;

						case "RIGHT":// �e���V�k����(�[�ݵe�����褺�e)
							if (targetElementP.x + targetElementS.width < ScrollBarP.x) {// �Y�j�M���󪺳̤jx�y�Фp����bx�y�СA��ܷj�M�������UI�Q���b�B��
								driver[j].swipe(ScrollBarP.x + errorX, targetElementP.y,
										ScrollBarP.x + ScrollBarS.width - errorX, targetElementP.y, 1000);
							} else if (targetElementP.x == ScrollBarP.x) {// �Y�j�M����x�y�е�����bx�y�СA�i���ܷj�M���󪺳���UI�Q���b�B��
								driver[j].swipe(targetElementP.x + targetElementS.width + errorX, targetElementP.y,
										ScrollBarP.x + ScrollBarS.width - errorX, targetElementP.y, 1000);
							}
							break;
						}

						wait[j].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appInputXpath))).click();
						// driver[j].findElement(By.xpath(appInputXpath)).click();
						break;
					}

					if (i == targetlist.size() - 1) {// �Ytargetlist���̫�@����Ƥ�粒��A�h�i��Srcoll�즲

						switch (scroll.toString()) {

						case "DOWN":
							driver[j].swipe(ScrollBarP.x + errorX, ScrollBarP.y + ScrollBarS.height - errorY,
									ScrollBarP.x + errorX, ScrollBarP.y + errorY, 1000);// �V�U����
							break;

						case "UP":
							driver[j].swipe(ScrollBarP.x + errorX, ScrollBarP.y + errorY, ScrollBarP.x + errorX,
									ScrollBarP.y + ScrollBarS.height - errorY, 1000);// �V�W����
							break;

						case "LEFT":
							driver[j].swipe(ScrollBarP.x + ScrollBarS.width - errorX, ScrollBarP.y + errorY,
									ScrollBarP.x + errorX, ScrollBarP.y + errorY, 1000);// �e���V������(�[�ݵe���k�褺�e)
							break;

						case "RIGHT":
							driver[j].swipe(ScrollBarP.x + errorX, ScrollBarP.y + errorY,
									ScrollBarP.x + ScrollBarS.width - errorX, ScrollBarP.y + errorY, 1000);// �e���V�k����(�[�ݵe�����褺�e)
							break;

						}
						SearchNumber++;// �֭p�j�M����
						targetlist.clear();// �M��targetlist
						targetlist = wait[j]
								.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(appElemntarray)));// �A�����o�stargetlist

						if (SearchNumber == 10) {// �j�M10�����S��줸��A�h���Xfor
							System.out.println("Can't find " + appInput);// �L�X�䤣��
							break;// ���Xfor
						} else {
							i = -1;// �YSearchNumber!=10�A�h�Oi=-1(�ت��G�A������for)
						}
					}
				}
				ErrorList[j] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception w) {
				System.out.print("[Error] Can't find " + appElemnt);
				System.out.print(" or [Error] can not find " + appElemntarray);
				System.out.println(" or [Error] can not find " + appInputXpath);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
			}
		}

	}
	/*
	 * �W�U�H���ư�n�� public void Swipe() { Random rand = new Random(); boolean
	 * items[] = { true, false }; for (int i = 0; i < driver.length; i++) { for
	 * (int j = 0; j < iterative; j++) { if (items[rand.nextInt(items.length)])
	 * { driver[i].swipe(startx, starty, endx, endy, 500); }else{
	 * driver[i].swipe(endx, endy, startx , starty , 500); } } } }
	 * 
	 */

	public void SubMethod_Result(boolean ErrorResult[], boolean result[]) {
		// �}��Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
		} catch (Exception e) {
			System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
		}
		for (int i = 0; i < driver.length; i++) {

			if (TestCase.DeviceInformation.deviceName.get(i).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
				char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
				TestCase.DeviceInformation.deviceName.get(i).toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
				Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// �ھ�NewUdid�A���w�Y�x�˸m��TestReport
																					// sheet
			} else {
				Sheet = workBook.getSheet(TestCase.DeviceInformation.deviceName.get(i).toString() + "_TestReport");// ���w�Y�x�˸m��TestReport
																													// sheet
			}

			if (ErrorResult[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error!!");
			} else if (result[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
			} else if (result[i] == false) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
			}
		}
		// ����g�JExcel�᪺�s�ɰʧ@
		try {
			FileOutputStream out = new FileOutputStream(new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
			workBook.write(out);
			out.close();
			workBook.close();
		} catch (Exception e) {
			System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
		}
	}
}
