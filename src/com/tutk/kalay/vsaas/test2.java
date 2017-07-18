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

public class test2 {

	public static void main(String[] args) throws InterruptedException {

		SeeTestAndroidDriver driver = null;
		DesiredCapabilities cap = new DesiredCapabilities();

		cap.setCapability(SeeTestCapabilityType.WAIT_FOR_DEVICE_TIMEOUT_MILLIS, 60 * 1000);
		cap.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 1);
		cap.setCapability(MobileCapabilityType.UDID, "CB5A210F3S");
		cap.setCapability(AndroidMobileCapabilityType.APP_PACKAGE, "com.lafresh.smile2");
		cap.setCapability(AndroidMobileCapabilityType.APP_ACTIVITY, "com.chinsoft.raise.screen.home.SplashActivity");

		try {
			driver = new SeeTestAndroidDriver(new URL("http://localhost:" + "4723" + "/wd/hub"), cap);
		} catch (Exception e) {
			;
		}
		WebDriverWait wait = new WebDriverWait(driver, 5);
		driver.findElement(By.xpath("//*[@resource-id='com.lafresh.smile2:id/records']")).click();

		String Scroll_list = "//*[@id='grid']";

		boolean state = true;
		while (state) {
			try {
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@text='CONVERSE']")));
				driver.findElement(By.xpath("//*[@text='CONVERSE']")).click();
				state = false;

				// wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@text='台中
				// 逢甲adidas 專賣店']")));

			} catch (Exception d) {

				Point p;// 元件座標
				Dimension s;// 元件大小
				WebElement e;
				e = driver.findElement(By.xpath(Scroll_list));// Scroll list
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.width * 0.1);
				int errorY = (int) Math.round(s.height * 0.1);
				driver.swipe(p.x + errorX, s.height - errorY, p.x + errorX, p.y + errorY, 1000);// 向下拖曳

			}
		}
	}

}
