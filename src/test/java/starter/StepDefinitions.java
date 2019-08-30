package starter;

import cucumber.api.java.en.Given;
import cucumber.api.java.en.Then;
import cucumber.api.java.en.When;
import net.thucydides.core.annotations.Steps;
import starter.navigation.DuckDuckGoHomePage;

public class StepDefinitions {
    @Steps
    DuckDuckGoHomePage home;

    @Given("^I am thirsty$")
    public void i_am_thirsty() throws Throwable {
    }

    @When("^I order a cappuccino$")
    public void i_order_a_cappuccino() throws Throwable {
    }

    @Then("^I should not receive a latte$")
    public void i_should_not_receive_a_latte() throws Throwable {
    }
//
//    @Given("^Sergey is on the DuckDuckGo home page$")
//    public void sergey_is_on_the_DuckDuckGo_home_page() {
//        // Write code here that turns the phrase above into concrete actions
//        home.me();
//    }
//
//    @When("^he searches for \"([^\"]*)\"$")
//    public void he_searches_for(String arg1) {
//        // Write code here that turns the phrase above into concrete actions
//        home.me();
//    }
//
//    @Then("^all the result titles should contain the word \"([^\"]*)\"$")
//    public void all_the_result_titles_should_contain_the_word(String arg1) {
//        // Write code here that turns the phrase above into concrete actions
//        home.me();
//    }

}
