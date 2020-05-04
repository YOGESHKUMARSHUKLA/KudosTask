package starter.navigation;

import net.serenitybdd.core.pages.PageObject;
import net.serenitybdd.core.pages.WebElementFacade;
import org.assertj.core.api.SoftAssertions;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.FindBy;

import java.time.temporal.ChronoUnit;
import java.util.concurrent.TimeUnit;

//@DefaultUrl("http://kudos-interview.s3-website-us-east-1.amazonaws.com/")
public class KudoLogin extends PageObject {

    @FindBy(xpath = "//input[@name='uname']")
    private WebElementFacade username;
    @FindBy(xpath = "//input[@name='password']")
    private WebElementFacade password;
    @FindBy(xpath = "//input[@class='submit']")
    private WebElementFacade loginBtn;
    @FindBy(xpath = "//p[@class='post-discription']")
    private WebElementFacade viewPost;
    @FindBy(xpath = "//span[@class='like-post']/img")
    private WebElementFacade likePost;
    @FindBy(xpath = "//textarea[@id='postArea']")
    private WebElementFacade createPost;
    @FindBy(xpath = "//input[@name='submit']")
    private WebElementFacade submitPost;



    public KudoLogin(WebDriver driver){super(driver);
    }

    public  void initialize(){
        getDriver().manage().deleteAllCookies();
        getDriver().manage().window().maximize();
        getDriver().manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
    }

    public  void inputUsername(String usernamee){
        username.sendKeys(usernamee);
    }
    public  void inputPass(String pass){
        password.sendKeys(pass);
    }
    public  void clickLogin(){
        loginBtn.click();
    }
    public void assertScreen(){
        withTimeoutOf(60, ChronoUnit.SECONDS).waitFor(createPost);
    }

    public void viewPost(){
        SoftAssertions softly = new SoftAssertions();
        softly.assertThat(viewPost.isDisplayed()).as("Nothing to View or Unable to View Post").isEqualTo(true);
    }

    public  void clickLike(){
        likePost.click();
    }

    public void createPOst(){
        createPost.sendKeys("Post");
        submitPost.click();
    }
}
