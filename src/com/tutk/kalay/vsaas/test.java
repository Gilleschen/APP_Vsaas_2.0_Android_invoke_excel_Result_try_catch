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
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("xpath=//*[@text='adidas-服飾']")));
		driver.findElement(By.xpath("xpath=//*[@text='adidas-服飾']")).click();

		String targetText = "AIR JORDAN 12 RETRO LOW BG";
		String targetXpath = "xpath=//*[@text='AIR JORDAN 12 RETRO LOW BG']";
		String Scroll_list = "xpath=//*[@id='pager']";
		String targetarray = "xpath=//*[@resource-id='com.lafresh.smile2:id/name']";
		String scroll = "RIGHT";
		List<WebElement> targetlist = driver.findElementsByXPath(targetarray);// 要搜尋的多筆類似元件清單
		String lastelement = targetlist.get(targetlist.size() - 1).getText().toString();// 要搜尋的多筆類似元件清單中，最後一筆字串(目的：判斷是否所有資料都搜尋過)

		for (int i = 0; i < targetlist.size(); i++) {

			if ((targetlist.get(i).getText().toString()).equals(targetText)) {// 若targetelement在targetlist清單中，則點擊targetelement
				WebElement ScrollBar, targetElement;
				Point ScrollBarP, targetElementP;// 元件座標
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

			if (i == targetlist.size() - 1) {// 若targetlist中最後一筆資料比對完後，則進行Srcoll拖曳

				Point p;// 元件座標
				Dimension s;// 元件大小
				WebElement e;
				e = driver.findElement(By.xpath(Scroll_list));// Scroll bar
				s = e.getSize();
				p = e.getLocation();
				int errorX = (int) Math.round(s.width * 0.1);
				int errorY = (int) Math.round(s.height * 0.1);

				if (scroll.equals("DOWN")) {
					driver.swipe(p.x + errorX, s.height - errorY, p.x + errorX, p.y + errorY, 1000);// 向下捲動
				} else if (scroll.equals("UP")) {
					driver.swipe(p.x + errorX, p.y + errorY, p.x + errorX, s.height - errorY, 1000);// 向上捲動
				} else if (scroll.equals("LEFT")) {
					driver.swipe(s.width - errorX, p.y + errorY, p.x + errorX, p.y + errorY, 1000);// 畫面向左捲動(觀看畫面右方內容)
				} else if (scroll.equals("RIGHT")) {
					driver.swipe(p.x + errorX, p.y + errorY, s.width - errorX, p.y + errorY, 1000);// 畫面向右捲動(觀看畫面左方內容)
				}
				Thread.sleep(1500);// 緩衝畫面更新後，在取得新targetlist
				targetlist.clear();// 清除targetlist
				targetlist = driver.findElementsByXPath(targetarray);// 再次取得新targetlist
				
				if (lastelement.equals(targetlist.get(targetlist.size() - 1).getText().toString())) {// 判斷新targetlist最後一筆資料是否與lastelement相同
					System.out.println("找不到" + targetText);// 若相同，表示Srcoll已拖曳至最後，則印出找不到，跳出for
					break;
				} else {// 若不相同，則更新lastelement最後筆資料
					lastelement = targetlist.get(targetlist.size() - 1).getText().toString();
				}
				i = -1;// 令i=-1(目的：再次執行for)
			}
		}

	}

}
