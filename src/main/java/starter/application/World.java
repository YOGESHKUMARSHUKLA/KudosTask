package starter.application;

import cucumber.api.Scenario;
import net.serenitybdd.core.Serenity;
import net.thucydides.core.util.EnvironmentVariables;
import net.thucydides.core.util.SystemEnvironmentVariables;

public class World {
    private static final String SESSION_VARIABLE_NAME = "World";
    private Order order;
    private Scenario scenario;

    public Scenario getScenario() {
        return scenario;
    }

    public void setScenario(Scenario scenario) {
        this.scenario = scenario;
    }

    public Order getOrder() {
        return order;
    }

    public void setOrder(Order order) {
        this.order = order;
    }

    public static World getWorld(){
        if (Serenity.hasASessionVariableCalled(SESSION_VARIABLE_NAME)){
            return Serenity.sessionVariableCalled(SESSION_VARIABLE_NAME);
        }
        else{
            World world =new World();
            world.setOrder(new Order(getEnvironmentVariable("orderNumber")));
            Serenity.setSessionVariable(SESSION_VARIABLE_NAME).to(world);
            return world;
        }
    }
    public static String getEnvironmentVariable(String key){
        EnvironmentVariables environmentVariables = SystemEnvironmentVariables.createEnvironmentVariables();
        return environmentVariables.getProperty(key);
    }
}
