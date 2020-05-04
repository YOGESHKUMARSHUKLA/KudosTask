package starter.Steps;

import net.thucydides.core.annotations.Step;
import starter.navigation.KudoLogin;

public class KudosSteps {
    KudoLogin kudoLogin;
    @Step
    public void launchKudos(){
        System.setProperty("webdriver.chrome.driver","D:/chromedriver.exe");
        kudoLogin.initialize();
        kudoLogin.setDefaultBaseUrl("http://kudos-interview.s3-website-us-east-1.amazonaws.com");
        kudoLogin.open();
    }
    @Step
    public void loginKudos(String username,String password){
        kudoLogin.inputUsername(username);
        kudoLogin.inputPass(password);
        kudoLogin.clickLogin();
        kudoLogin.assertScreen();
    }

    @Step
    public void viewPost(){
        kudoLogin.viewPost();
    }

    @Step
    public void likePost(){
        kudoLogin.clickLike();
    }
    @Step
    public void createPost(){
        kudoLogin.createPOst();
    }
}
