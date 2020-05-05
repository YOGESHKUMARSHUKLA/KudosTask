package starter.generic;

import cucumber.api.Scenario;
import net.serenitybdd.core.Serenity;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import starter.application.World;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Evidence {
    private Evidence(){throw new IllegalStateException("Utility Class");}

    public static void takeScreenshot(String title, WebDriver driver) throws IOException {
        TakesScreenshot takesScreenshots = (TakesScreenshot) driver;
        File screenshot = takesScreenshots.getScreenshotAs(OutputType.FILE);
        Serenity.recordReportData().asEvidence().withTitle(title).downloadable().fromFile(screenshot.toPath());
        embedToCucumberReport(Files.readAllBytes(screenshot.toPath()),"image/png");
    }

    public static void recordTextEvidence(String title, String content) {
        Serenity.recordReportData().asEvidence().withTitle(title).andContents(content);
        embedToCucumberReport(content.getBytes(),"text/plain");
    }

    public static void recordFileEvidence(String title, byte[] content,String extension,String mimeType) throws IOException {
        Path path = Paths.get("resources//tempfile//temp."+extension);
        Files.write(path,content);
        Serenity.recordReportData().asEvidence().withTitle(title).downloadable().fromFile(path);
        embedToCucumberReport(content,mimeType);
    }

    private static void embedToCucumberReport(byte[] content,String mimeType) {
        Scenario scenario = World.getWorld().getScenario();
        if(scenario!=null){
            scenario.embed(content,mimeType);
        }
    }

    public static void recordHTMLEvidence(String title,String contenet) throws IOException {
        if(!contenet.startsWith("<html>" )|| !contenet.startsWith("<!DOCTYPE")){
            contenet="<html><head></head><body>"+contenet+"</body></html>";
        }
        recordFileEvidence(title,contenet.getBytes(),"html","text/html");
    }

    public static void recordPDFEvidence(String title,byte[] contenet) throws IOException {
        recordFileEvidence(title,contenet,"pdf","application/pdf");
    }

    public static void recordXLSXEvidence(String title,byte[] contenet) throws IOException {
        recordFileEvidence(title,contenet,"xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    public static void recordXLSEvidence(String title,byte[] contenet) throws IOException {
        recordFileEvidence(title,contenet,"xls","application/vnd.ms-excel");
    }
    public static void recordtextFileEvidence(String title,byte[] contenet) throws IOException {
        recordFileEvidence(title,contenet,"txt","text");
    }
}
