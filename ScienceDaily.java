package com.parallelexecution;

import java.io.File;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;


public class ScienceDaily {

	    static WebDriver driver;
	
	    public void writeExcel() throws Exception {
		String strsciencedailyJournalR = null;
		//String strsciencedailyJournalR1 = null;
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Krishna\\Desktop\\Selenium\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		//driver.get("https://www.sciencedaily.com/news/top/health/");
		driver.get("https://www.sciencedaily.com/news/health_medicine/");
		//driver.get("https://www.sciencedaily.com/news/top/");
		Thread.sleep(1000);
		WebElement headlines = driver.findElement(By.xpath("//a[text()='Headlines']"));
		headlines.click();
    	//act.moveToElement(driver.findElement(By.xpath("((//ul[@class='fa-ul list-padded'])[1])/li[1]"))).click();
		//System.out.println("before");
		WebElement datewise = driver.findElement(By.xpath("//div[@class='headlines-date'][1]"));
		//System.out.println(datewise.getText());
		WebElement elem = driver.findElement(By.xpath("(//div[@class='headlines-date'])[1]/following-sibling::ul[1]"));
		//WebElement elem = driver.findElement(By.xpath("(//div[@class='headlines-date'])[2]/following-sibling::ul[1]"));
		//WebElement elem = driver.findElement(By.xpath("(//div[@class='headlines-date'])[1]/following-sibling::ul[1]"));
		//System.out.println(elem.getText());
		//elem.click();
		//System.out.println("after");
		Thread.sleep(1000);
		
		// This data needs to be written (Object[])
		Map<String, Object[]> cmninfo = new TreeMap<String, Object[]>();
		cmninfo.put("1", new Object[] { "Title", "Description","Keywords", "Content", "References","Materials",
				"ImageDescription", "HealthCenters","First Sentence of Article","Source" });
		
		
		 int intsize = driver.findElements(By.xpath("((//div[@class='headlines-date'])[1]/following-sibling::ul[1])/li")).size();
		//int intsize = driver.findElements(By.xpath("((//div[@class='headlines-date'])[2]/following-sibling::ul[1])/li")).size();
		System.out.println(intsize);
		for (int i = 1; i <intsize+1; i++) 
		{
			Thread.sleep(3000);
		    WebElement TopHealthTopics = driver.findElement(By.xpath("((//div[@class='headlines-date'])[1]/following-sibling::ul[1])/li["+i+"]/a"));
			//WebElement TopHealthTopics = driver.findElement(By.xpath("((//div[@class='headlines-date'])[1]/following-sibling::ul[1])/li["+i+"]/a"));
			//WebElement TopHealthTopics = driver.findElement(By.xpath("((//div[@class='headlines-date'])[2]/following-sibling::ul[1])/li["+i+"]/a"));
			//System.out.println(TopHealthTopics.getText());
			//TopHealthTopics.click();
			Thread.sleep(2000);
			Actions action = new Actions(driver);
            action.moveToElement(TopHealthTopics).click().build().perform();			
			//WebElement TopHealthTopics = driver.findElement(By.xpath("//a[@data-tab='#science_tab_" + i + "']"));
			//TopHealthTopics.click();
            Thread.sleep(2000);
			WebElement HeadingTitle = driver.findElement(By.xpath("//h1[@class='headline']"));
			String strHeadingTitle = HeadingTitle.getText();
			WebElement StrSource = driver.findElement(By.xpath("//dd[@id='source']"));
			String source1 = StrSource.getText();
			//System.out.println(strHeadingTitle);
			WebElement Sciencedes = driver.findElement(By.xpath("//dd[@id='abstract']"));
			String strSciencedes = Sciencedes.getText();
			WebElement sciencedailyImageDescription = driver.findElement(By.xpath("//p[@id='first']"));
			String strsciencedailyImageDescription = sciencedailyImageDescription.getText();
			WebElement sciencedailyContent = driver.findElement(By.xpath("//div[@id='text']"));
			String strsciencedailyContent = sciencedailyContent.getText();
			String strFullContent = (strsciencedailyImageDescription+strsciencedailyContent);
			//System.out.println(strsciencedailyContent);	
			//WebElement sciencedailyJournalR = driver.findElement(By.xpath("//ol[@class='journal']"));
			//System.out.println("***************************test");
			WebElement sciencedailyMaterials = driver.findElement(By.xpath("//div[@id='story_source']/p[2]"));
			String strsciencedailyMaterials = sciencedailyMaterials.getText();

			try{
				 if(driver.findElement(By.xpath("//ol[@class='journal']")).isDisplayed()){
					 WebElement sciencedailyJournalR = driver.findElement(By.xpath("//ol[@class='journal']"));
						strsciencedailyJournalR = sciencedailyJournalR.getText();
		    			System.out.println("@@@@@@@@@@@@@@@@@@@@@@ is present"); 
		    			Thread.sleep(1000);
		    			}
				}catch(Exception e){
				Thread.sleep(1000);
				driver.findElement(By.xpath("//a[text()='APA']")).click();
				WebElement referenceAPA = driver.findElement(By.xpath("//div[@id='citation_apa']"));
				strsciencedailyJournalR = referenceAPA.getText();
				System.out.println(strsciencedailyJournalR);
				String replacedtext = strsciencedailyJournalR.toString();
				System.out.println(replacedtext);
				//System.out.println(strsciencedailyJournalR1.replace('i', 'E'));
												
				}
									
		//	WebElement sciencedailyMaterials = driver.findElement(By.xpath("//div[@id='story_source']/p[2]"));
			//String strsciencedailyMaterials = sciencedailyMaterials.getText();

			
			//System.out.println(strsciencedailyMaterials);

			WebElement sciencedailyImageDescription1 = driver.findElement(By.xpath("//p[@id='first']"));

			String FirstSentenceofArticle1 = sciencedailyImageDescription1.getText();

			//System.out.println(strsciencedailyImageDescription);
			

			// Create blank workbook
			
			@SuppressWarnings("resource")
			XSSFWorkbook workbook = new XSSFWorkbook();

			// Create a blank sheet

			XSSFSheet spreadsheet = workbook.createSheet(" CMN ");

			// Create row object

			XSSFRow row;

			int count = i + 2;

			String strcount = Integer.toString(count);
			
					cmninfo.put(strcount, new Object[] {

					strHeadingTitle, strSciencedes, "", strFullContent, strsciencedailyJournalR, strsciencedailyMaterials, "", "Current Medical News",FirstSentenceofArticle1,source1});
			 
			// Iterate over data and write to sheet

			Set<String> keyid = cmninfo.keySet();
			int rowid = 0;
			for (String key : keyid) {
				row = spreadsheet.createRow(rowid++);
				Object[] objectArr = cmninfo.get(key);
				int cellid = 0;
				for (Object obj : objectArr) {
					XSSFCell cell = row.createCell(cellid++);
					cell.setCellValue((String) obj);
				}

			}
			

			// Write the workbook in file system

			FileOutputStream out = new FileOutputStream(

					new File("C:\\Users\\Krishna\\Desktop\\CMN\\TopHealth.xlsx"));

			workbook.write(out);

			out.close();

			System.out.println("TopHealthNews written successfully");

			Thread.sleep(2000);

			driver.findElement(By.xpath("//a[@class='navbar-brand dropdown-toggle']")).click();
			//Select Sciencedaily = new Select(driver.findElement(By.xpath("//a[@class='navbar-brand dropdown-toggle']")));
			//Select Sciencedaily = new Select (driver.findElement(By.xpath("//ul[@class='dropdown-menu brand']")));
			//Sciencedaily.selectByVisibleText("Latest News");
			WebElement Sciencedaily = driver.findElement(By.xpath("(//a[text()='Latest News'])[1]"));
			Sciencedaily.click();
			Thread.sleep(2000);
		    WebElement topHealth = driver.findElement(By.xpath("//a[text()='Top Health']"));
			//WebElement topHealth = driver.findElement(By.xpath("//a[text()='TOP NEWS']"));
			action.moveToElement(topHealth).click().build().perform();
		    //topHealth.click();
			WebElement headlines1 = driver.findElement(By.xpath("//a[text()='Headlines']"));
			action.moveToElement(headlines1).click().build().perform();
			//headlines1.click();		
			

		}

	}

	public static void main(String[] args) throws Exception {
		ScienceDaily scienceDaily = new ScienceDaily();
		scienceDaily.writeExcel();
		driver.close();

	}

}
