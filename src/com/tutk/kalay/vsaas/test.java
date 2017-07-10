package com.tutk.kalay.vsaas;

import java.net.URL;
import java.util.List;
import java.util.Random;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.experitest.appium.SeeTestAndroidDriver;
import com.experitest.appium.SeeTestCapabilityType;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class test {

	public static void main(String[] args) throws InterruptedException {
		Random rand = new Random();
		SeeTestAndroidDriver driver = null;
		DesiredCapabilities cap = new DesiredCapabilities();
		WebDriverWait wait = null;

		cap.setCapability(SeeTestCapabilityType.WAIT_FOR_DEVICE_TIMEOUT_MILLIS, 60 * 1000);
		cap.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 30);
		cap.setCapability(MobileCapabilityType.UDID, "CB5A210F3S");
		cap.setCapability(AndroidMobileCapabilityType.APP_PACKAGE, "com.tutk.kalayvsaas");
		cap.setCapability(AndroidMobileCapabilityType.APP_ACTIVITY, "com.tutk.kalayvsaas.SplashScreenActivity");

		try {
			driver = new SeeTestAndroidDriver(new URL("http://localhost:" + "4723" + "/wd/hub"), cap);
		} catch (Exception e) {
			;
		}

		// String target = "wuunet2";
		wait = new WebDriverWait(driver, 30);
		Thread.sleep(6000);

		List<WebElement> tvnickname = driver
				.findElementsByXPath("//*[@resource-id='com.tutk.kalayvsaas:id/tvNickname']");
		List<WebElement> panel = driver.findElementsByXPath("//*[@resource-id='com.tutk.kalayvsaas:id/layoutPanel']");
		List<WebElement> tvconnect = driver
				.findElementsByXPath("//*[@resource-id='com.tutk.kalayvsaas:id/tvConnection']");
		for (int i = 0; i < tvnickname.size(); i++) {
			if ((tvnickname.get(i).getText().toString()).equals("E8")) {
				Thread.sleep(6000);
				WebElement target = panel.get(i);
				target.click();
				break;

			}
			if (i == tvnickname.size() - 1) {

				Point p;// 元件座標
				Dimension s;// 元件大小
				WebElement e;
				e = driver.findElement(By.xpath("//*[@id='layoutCamLst']"));
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.width * 0.01);
				int errorY = (int) Math.round(s.height * 0.01);

				driver.swipe(p.x + errorX, s.height - errorY, p.x + errorX, p.y + errorY, 1000);

				tvnickname.clear();
				panel.clear();
				tvconnect.clear();
				tvnickname = driver.findElementsByXPath("//*[@resource-id='com.tutk.kalayvsaas:id/tvNickname']");
				panel = driver.findElementsByXPath("//*[@resource-id='com.tutk.kalayvsaas:id/layoutPanel']");
				tvconnect = driver.findElementsByXPath("//*[@resource-id='com.tutk.kalayvsaas:id/tvConnection']");
				i = -1;
			}
		}

		// driver.swipe(p.x + errorX, s.height - errorY, p.x + errorX, p.y +
		// errorY, 1000);

		// driver.swipe(p.x + errorX, p.y + errorY, p.x + errorX, s.height -
		// errorY, 1000);

	}

}
