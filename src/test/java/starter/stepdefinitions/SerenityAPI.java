package starter.stepdefinitions;

import cucumber.api.java.en.Given;
import io.restassured.response.Response;
import net.serenitybdd.rest.SerenityRest;

public class SerenityAPI {
    private Response response;

    @Given("^the user is connected to API \"([^\"]*)\"$")
    public void the_user_is_connected_to_API(String arg1) {
    // Write code here that turns the phrase above into concrete actions
        response= SerenityRest.when().get(arg1);
        System.out.println(response.getBody().asString());

}

}
