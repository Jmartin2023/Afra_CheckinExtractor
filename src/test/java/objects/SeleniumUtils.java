package objects;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeDriverService;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.chromium.ChromiumDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

import java.io.File;
import java.time.Duration;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class SeleniumUtils {

    private WebDriver driver;
    private WebDriverWait wait5, wait10, wait20;
    private String downloadPath = System.getProperty("user.dir");


    public SeleniumUtils() {
        this.driver = initiateDriver(this.downloadPath);
        intiateWait();
    }

    public SeleniumUtils(String downloadPath) {
        this.downloadPath = downloadPath;
        this.driver = initiateDriver(downloadPath);
        intiateWait();
    }
    
    public SeleniumUtils(WebDriver driver) {
        this.driver = driver;
        intiateWait();
    }
    
    public WebDriver getDriver() {
    	return this.driver;
    }

    public WebDriverWait getWait10() {
        return wait10;
    }

    public WebDriverWait getWait20() {
        return wait20;
    }

    public WebDriver initiateDriver(String path) {

        HashMap<String, Object> chromePrefs = new HashMap<>();
      
        chromePrefs.put("download.prompt_for_download", false);
        chromePrefs.put("plugins.always_open_pdf_externally", true); // don't open viewer, download instea
   //    chromePrefs.put("profile.default_content_settings.popups", 0);
        chromePrefs.put("download.default_directory", path);
    
        chromePrefs.put("printing.print_preview_sticky_settings.appState",
                "{\"recentDestinations\":[{\"id\":\"Save as PDF\",\"origin\":\"local\",\"account\":\"\"}],\"selectedDestinationId\":\"Save as PDF\",\"version\":2}");
        chromePrefs.put("savefile.default_directory", System.getProperty("user.dir")+"\\DownloadedFiles");
     
        
        
        ChromeOptions options = new ChromeOptions();

        options.setCapability(CapabilityType.ACCEPT_INSECURE_CERTS, true);
        options.setCapability(ChromeOptions.CAPABILITY, options);
     //   options.setHeadless(false);
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--ignore-ssl-errors=yes");
        options.addArguments("--ignore-certificate-errors");
        
        //adding this for profile saving
   //     options.addArguments("--user-data-dir=D:\\chrome-profiles");
     //   options.addArguments("--profile-directory=Profile 2");
        
      
  //      options.addArguments("--disable-print-preview");
   //     options.addArguments("--kiosk-printing");
        options.setExperimentalOption("prefs", chromePrefs);
        
  //     options.addExtensions(new File("C:\\Users\\jmartin\\Downloads\\Authenticator.crx"));
      
     
        options.setExperimentalOption("excludeSwitches", Collections.singletonList("enable-automation"));

        System.setProperty(ChromeDriverService.CHROME_DRIVER_SILENT_OUTPUT_PROPERTY, "true");
 //       WebDriverManager.chromedriver().setup();
     System.setProperty("webdriver.chrome.driver", "\\\\10.170.193.71\\Development Team\\RPA\\Development\\chromedriver\\chromedriver-win64\\chromedriver152.exe");
    //    System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"\\chromedriver.exe");
//        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver(options);
      

       Map<String, Object> params = new HashMap<>();
       params.put("source", "Object.defineProperty(navigator, 'webdriver', { get: () => undefined})");

        ((ChromiumDriver) driver).executeCdpCommand("Page.addScriptToEvaluateOnNewDocument", params);

        driver.manage().window().maximize();

        System.out.println("Chrome driver started successfully");
        
        return driver;

    }

    public void pauseVisibility(WebElement element, int sec) {
    	try {
    		WebDriverWait wait = getWait(sec);
    		wait.until(ExpectedConditions.visibilityOf(element));
    	} catch (Exception e) {
    		System.out.println(e.getMessage());
    	}
    }

    public void pauseInvisibility(WebElement element, int sec) {
    	try {
	        WebDriverWait wait = getWait(sec);
	        wait.until(ExpectedConditions.invisibilityOf(element));
    	} catch (Exception e) {
			System.out.println(e.getMessage());
		}
    }

    public void pauseClick(WebElement element, int sec) {
    	try {
			WebDriverWait wait = getWait(sec);
			wait.until(ExpectedConditions.elementToBeClickable(element));
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
    }

    public void pauseWindowsCount(WebElement element, int sec, int numOfWindows) {
    	try {
	    	WebDriverWait wait = getWait(sec);
	        wait.until(ExpectedConditions.numberOfWindowsToBe(numOfWindows));
    	} catch (Exception e) {
			System.out.println(e.getMessage());
		}
    }
    
    private void intiateWait() {
        this.wait5 = new WebDriverWait(this.driver, Duration.ofSeconds(5));
        this.wait10 = new WebDriverWait(this.driver, Duration.ofSeconds(10));
        this.wait20 = new WebDriverWait(this.driver, Duration.ofSeconds(20));
    }
    
    private WebDriverWait getWait(int sec) {
		WebDriverWait wait;
		if(sec==5) wait = this.wait5;
        else if(sec==10) wait = this.wait10;
        else if(sec==20) wait = this.wait20;
        else wait = new WebDriverWait(this.driver, Duration.ofSeconds(sec));
		return wait;
	}

}
