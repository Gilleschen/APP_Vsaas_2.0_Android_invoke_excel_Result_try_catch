package com.tutk.kalay.vsaas;

import java.net.URL;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.experitest.appium.SeeTestAndroidDriver;
import com.experitest.appium.SeeTestCapabilityType;
import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class test {

	public static void main(String[] args) throws InterruptedException {

		SeeTestAndroidDriver driver = null;
		DesiredCapabilities cap = new DesiredCapabilities();

		cap.setCapability(SeeTestCapabilityType.WAIT_FOR_DEVICE_TIMEOUT_MILLIS, 60 * 1000);
		cap.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 10);
		cap.setCapability(MobileCapabilityType.UDID, "HT4C4YC01143");
		cap.setCapability(AndroidMobileCapabilityType.APP_PACKAGE, "com.lafresh.smile2");
		cap.setCapability(AndroidMobileCapabilityType.APP_ACTIVITY, "com.chinsoft.raise.screen.home.SplashActivity");

		try {
			driver = new SeeTestAndroidDriver(new URL("http://localhost:" + "4723" + "/wd/hub"), cap);
		} catch (Exception e) {
			;
		}
		WebDriverWait wait = new WebDriverWait(driver, 5);
		driver.findElement(By.xpath("xpath=//*[@resource-id='com.lafresh.smile2:id/records']")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("xpath=//*[@text='adidas-�A��']")));
		driver.findElement(By.xpath("xpath=//*[@text='adidas-�A��']")).click();

		String targetText = "AIR JORDAN 12 RETRO LOW BG";
		String targetXpath = "xpath=//*[@text='AIR JORDAN 12 RETRO LOW BG']";
		String Scroll_list = "xpath=//*[@id='pager']";
		String targetarray = "xpath=//*[@resource-id='com.lafresh.smile2:id/name']";
		String scroll = "RIGHT";
		List<WebElement> targetlist = driver.findElementsByXPath(targetarray);// �n�j�M���h����������M��
		String lastelement = targetlist.get(targetlist.size() - 1).getText().toString();// �n�j�M���h����������M�椤�A�̫�@���r��(�ت��G�P�_�O�_�Ҧ���Ƴ��j�M�L)

		for (int i = 0; i < targetlist.size(); i++) {

			if ((targetlist.get(i).getText().toString()).equals(targetText)) {// �Ytargetelement�btargetlist�M�椤�A�h�I��targetelement
				WebElement ScrollBar, targetElement;
				Point ScrollBarP, targetElementP;// ����y��
				Dimension ScrollBarS, targetElementS;

				ScrollBar = driver.findElement(By.xpath(Scroll_list));// Scroll
																		// bar
				targetElement = driver.findElement(By.xpath(targetXpath));

				ScrollBarP = ScrollBar.getLocation();
				targetElementP = targetElement.getLocation();

				ScrollBarS = ScrollBar.getSize();
				targetElementS = targetElement.getSize();

				int errory = (int) Math.round(ScrollBarS.height * 0.1);
				int errorx = (int) Math.round(ScrollBarS.width * 0.1);

				if (scroll.equals("DOWN")) {
					if (targetElementS.height < 0 || targetElementP.y > ScrollBarS.height) {
						driver.swipe(targetElementP.x, ScrollBarS.height + ScrollBarP.y - errory, targetElementP.x,
								ScrollBarP.y, 1000);
					} else {
						driver.swipe(targetElementP.x, targetElementP.y, targetElementP.x, ScrollBarP.y, 1000);
					}
				} else if (scroll.equals("UP")) {
					if (targetElementP.y < 0 || targetElementP.y < ScrollBarP.y) {
						driver.swipe(targetElementP.x, ScrollBarP.y, targetElementP.x,
								ScrollBarS.height + ScrollBarP.y - errory, 1000);
					} else {
						driver.swipe(targetElementP.x, ScrollBarP.y, targetElementP.x,
								ScrollBarP.y + ScrollBarS.height - errory, 1000);
					}
				} else if (scroll.equals("LEFT")) {
					if (targetElementP.x > ScrollBarS.width) {
						driver.swipe(ScrollBarS.width - errorx, targetElementP.y, ScrollBarP.x + errorx,
								targetElementP.y, 1000);
					}
				} else if (scroll.equals("RIGHT")) {

					if (targetElementP.x < 0) {
						driver.swipe(ScrollBarP.x + errorx, targetElementP.y, ScrollBarS.width - errorx,
								targetElementP.y, 1000);
					}
				}

				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(targetXpath)));
				driver.findElement(By.xpath(targetXpath)).click();
				break;
			}

			if (i == targetlist.size() - 1) {// �Ytargetlist���̫�@����Ƥ�粒��A�h�i��Srcoll�즲

				Point p;// ����y��
				Dimension s;// ����j�p
				WebElement e;
				e = driver.findElement(By.xpath(Scroll_list));// Scroll bar
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.width * 0.1);
				int errorY = (int) Math.round(s.height * 0.1);

				if (scroll.equals("DOWN")) {
					driver.swipe(p.x + errorX, s.height - errorY, p.x + errorX, p.y + errorY, 1000);// �V�U����
				} else if (scroll.equals("UP")) {
					driver.swipe(p.x + errorX, p.y + errorY, p.x + errorX, s.height - errorY, 1000);// �V�W����
				} else if (scroll.equals("LEFT")) {
					driver.swipe(s.width - errorX, p.y + errorY, p.x + errorX, p.y + errorY, 1000);// �e���V������(�[�ݵe���k�褺�e)
				} else if (scroll.equals("RIGHT")) {
					driver.swipe(p.x + errorX, p.y + errorY, s.width - errorX, p.y + errorY, 1000);// �e���V�k����(�[�ݵe�����褺�e)
				}
				Thread.sleep(1500);// �w�ĵe����s��A�b���o�stargetlist
				targetlist.clear();// �M��targetlist
				targetlist = driver.findElementsByXPath(targetarray);// �A�����o�stargetlist
				
				if (lastelement.equals(targetlist.get(targetlist.size() - 1).getText().toString())) {// �P�_�stargetlist�̫�@����ƬO�_�Plastelement�ۦP
					System.out.println("�䤣��" + targetText);// �Y�ۦP�A���Srcoll�w�즲�̫ܳ�A�h�L�X�䤣��A���Xfor
					break;
				} else {// �Y���ۦP�A�h��slastelement�̫ᵧ���
					lastelement = targetlist.get(targetlist.size() - 1).getText().toString();
				}
				i = -1;// �Oi=-1(�ت��G�A������for)
			}
		}

	}

}
