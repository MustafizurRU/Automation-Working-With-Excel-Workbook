import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;


public class SentimentAnalyzer {
	static WebDriver driver;
	static String projectpath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static String comment="";
	static String parent_comment="";
	static String combine_com_parent="";
	static int rowNumber;
	public static void main(String[] args) throws InterruptedException {
		
		try {
			
			
			projectpath = System.getProperty("user.dir");
			FileInputStream file = new FileInputStream(projectpath + "/ExcelWork/Demo_Semantic_labelling.xlsx");
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheet("Sheet1");
			rowNumber = sheet.getPhysicalNumberOfRows();
			for(int i=1;i<=rowNumber;i++)
			{
				comment = sheet.getRow(i).getCell(0).getStringCellValue();
				parent_comment = sheet.getRow(i).getCell(1).getStringCellValue();
				combine_com_parent = sheet.getRow(i).getCell(2).getStringCellValue();
				
				System.setProperty("webdriver.gecko.driver", "C:\\Dev\\Drivers\\geckodriver.exe");
				driver = new FirefoxDriver();
				driver.get("https://www.danielsoper.com/sentimentanalysis/default.aspx");
				Thread.sleep(1000);
				//driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
				driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				
				driver.findElement(By.xpath("/html/body/form/div[3]/div/div/table/tbody/tr[3]/td[1]/div[1]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/table/tbody/tr/td[2]/a")).click();
				Thread.sleep(1000);
				driver.findElement(By.id("accordionPaneSentimentAnalysis_content_txtText")).sendKeys(comment);
				Thread.sleep(1000);
			
				driver.findElement(By.xpath("//*[@id=\"accordionPaneSentimentAnalysis_content_btnAnalyzeText\"]")).click();
				Thread.sleep(1000);
				
				Double num = Double.parseDouble(driver.findElement(By.xpath("/html/body/form/div[3]/div/div/table/tbody/tr[3]/td[1]/div[1]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/span/span[2]")).getText());
				if(num>=0)
				{
				
					sheet.getRow(i).getCell(3).setCellValue(num);
					sheet.getRow(i).getCell(4).setCellValue(1);
				}else {
					sheet.getRow(i).getCell(3).setCellValue(num);
					sheet.getRow(i).getCell(4).setCellValue(0);
				}
				
				Thread.sleep(1000);
				
				driver.findElement(By.xpath("/html/body/form/div[3]/div/div/table/tbody/tr[3]/td[1]/div[1]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/table/tbody/tr/td[2]/a")).click();
				Thread.sleep(1000);
				driver.findElement(By.id("accordionPaneSentimentAnalysis_content_txtText")).sendKeys(parent_comment);
				Thread.sleep(1000);
			
				driver.findElement(By.xpath("//*[@id=\"accordionPaneSentimentAnalysis_content_btnAnalyzeText\"]")).click();
				Thread.sleep(1000);
				
				Double num1 = Double.parseDouble(driver.findElement(By.xpath("/html/body/form/div[3]/div/div/table/tbody/tr[3]/td[1]/div[1]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/span/span[2]")).getText());
				if(num1>=0)
				{
				
					sheet.getRow(i).getCell(5).setCellValue(num1);
					sheet.getRow(i).getCell(6).setCellValue(1);
				}else {
					sheet.getRow(i).getCell(5).setCellValue(num1);
					sheet.getRow(i).getCell(6).setCellValue(0);
				}
				Thread.sleep(1000);
				driver.findElement(By.xpath("/html/body/form/div[3]/div/div/table/tbody/tr[3]/td[1]/div[1]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/table/tbody/tr/td[2]/a")).click();
				Thread.sleep(1000);
				driver.findElement(By.id("accordionPaneSentimentAnalysis_content_txtText")).sendKeys(combine_com_parent);
				Thread.sleep(1000);
			
				driver.findElement(By.xpath("//*[@id=\"accordionPaneSentimentAnalysis_content_btnAnalyzeText\"]")).click();
				Thread.sleep(1000);
				
				Double num2 = Double.parseDouble(driver.findElement(By.xpath("/html/body/form/div[3]/div/div/table/tbody/tr[3]/td[1]/div[1]/div[2]/div/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/span/span[2]")).getText());
				if(num2>=0)
				{
				
					sheet.getRow(i).getCell(7).setCellValue(num2);
					sheet.getRow(i).getCell(8).setCellValue(1);
				}else {
					sheet.getRow(i).getCell(7).setCellValue(num2);
					sheet.getRow(i).getCell(8).setCellValue(0);
				}
				file.close();
				FileOutputStream fileOut = new FileOutputStream(projectpath + "/ExcelWork/Demo_Semantic_labelling.xlsx");
				workbook.write(fileOut);
				fileOut.close();
				driver.close();
				
				
			}
			
			
			
		}catch (Exception e) {
			// TODO: handle exception
			e.getMessage();
			e.getCause();
			e.getStackTrace();
		}
		System.out.println("Task completed fully");
		
	}
}
