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
import com.experitest.appium.SeeTestIOSDriver;

import io.appium.java_client.MultiTouchAction;
import io.appium.java_client.TouchAction;
import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.IOSMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class test {

	public static void main(String[] args) {
		boolean state = true;
		SeeTestIOSDriver driver = null;
		DesiredCapabilities cap = new DesiredCapabilities();
		WebDriverWait wait = null;

		cap.setCapability(SeeTestCapabilityType.WAIT_FOR_DEVICE_TIMEOUT_MILLIS, 60000);
		cap.setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, 30);
		cap.setCapability(MobileCapabilityType.UDID, "7e9529ce5c4ee1c7de7598e7ac26e25e2d1f8700");
		cap.setCapability(IOSMobileCapabilityType.BUNDLE_ID, "com.tutk.kalayvsaas.v3.jenkins");

		// cap.setCapability(MobileCapabilityType.UDID, "CB5A210F3S");
		// cap.setCapability(AndroidMobileCapabilityType.APP_PACKAGE,
		// "com.tutk.kalayvsaas");
		// cap.setCapability(AndroidMobileCapabilityType.APP_ACTIVITY,
		// "com.tutk.kalayvsaas.SplashScreenActivity");

		try {
			driver = new SeeTestIOSDriver(new URL("http://localhost:" + "4723" + "/wd/hub"), cap);
		} catch (Exception e) {
			;
		}

		wait = new WebDriverWait(driver, 5);
		driver.findElement(By.xpath("//*[@text='device btn setting n' and ./preceding-sibling::*[@text='已連線']]"))
				.click();

		Point p1, p2;// p1 為起點;p2為終點

		p1 = driver.findElement(By.xpath("//*[@text='格式化SD卡']")).getLocation();
		p2 = driver.findElement(By.xpath("//*[@class='UIAImage']")).getLocation();
		driver.swipe(p1.x, p1.y, p1.x, p1.y - (p1.y - p2.y), 1000);

		driver.findElement(By.xpath("//*[@text='無線網路']")).click();

		while (state) {
			List<WebElement> elements = driver.findElementsByXPath("//*[@XCElementType='XCUIElementTypeStaticText']");

			for (int i = 0; i < elements.size(); i++) {
				System.out.println(elements.get(i).getText());
				if (elements.get(i).getText().equals("Ccphone")) {
					try{
						wait.until(ExpectedConditions.elementToBeClickable(By.xpath(elements.get(i).toString())));
						driver.findElement(By.xpath(elements.get(i).toString())).click();
						state = false;
						break;
					}catch(Exception e){
						;
					}
					

				}

			}
			if (state) {

				Point p;// 元件座標
				Dimension s;// 元件大小
				WebElement e;
				e = driver.findElement(By.xpath("//*[@class='UIATable']"));
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.width * 0.01);
				int errorY = (int) Math.round(s.height * 0.01);
				driver.swipe(p.x + errorX, s.height - errorY, p.x + errorX, p.y + errorY, 1000);

			}
		}

	}

}
