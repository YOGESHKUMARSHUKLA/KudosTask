package starter.stepdefinitions;

import cucumber.api.java.en.And;
import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import net.thucydides.core.annotations.Steps;
import starter.Steps.KudosSteps;

public class KudosDefinition {

    @Steps
    KudosSteps kudosSteps;

    @Given("^the user is on Kudos Login Page$")
    public void theUserIsOnKudosLoginPage() {
        kudosSteps.launchKudos();

    }

    @When("^user login with \"([^\"]*)\" username and \"([^\"]*)\" password$")
    public void userLoginWithUsernameAndPassword(String username, String password) {
        kudosSteps.loginKudos(username,password);
    }

    @Then("^user view the post$")
    public void userViewThePost() {
        kudosSteps.viewPost();
    }

    @And("^user like the post$")
    public void userLikeThePost() {
        kudosSteps.likePost();
    }

    @When("^user create post$")
    public void userCreatePost() {
        kudosSteps.createPost();
    }
}
