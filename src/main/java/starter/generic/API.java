package starter.generic;


import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.http.Method;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import org.apache.log4j.Logger;
import org.hamcrest.Matcher;
import org.hamcrest.Matchers;

import java.util.Arrays;
import java.util.List;

import static io.restassured.config.EncoderConfig.encoderConfig;
import static net.serenitybdd.rest.SerenityRest.rest;

public class API {
    private static final String AUTHORIZATION = "Authorization";
    private static final Logger LOGGER = Logger.getLogger(API.class);
    protected static int maxRetryTime =1;
    protected int attempt =1;
    private static List<Integer> successCode;
    private static List<Integer> failRetryCode;

    static {
        successCode = Arrays.asList(200,204);
        failRetryCode =Arrays.asList(400,500,403);
    }
    private String domain;
    private String endPoint;
    private Response response;
    private RequestSpecification spec;
    protected String payload;

    public API (String domain,String endPoint){
        this.domain =domain;
        this.endPoint=endPoint;
        spec = rest().contentType(ContentType.XML).accept(ContentType.XML).config(RestAssured.config()
                .encoderConfig(encoderConfig().defaultContentCharset("UTF-8").appendDefaultContentCharsetToContentTypeIfUndefined(false)));
    }

    public static void init(){}

    public API (String domain,String endPoint,ContentType contentType){
        this.domain =domain;
        this.endPoint=endPoint;
        spec = rest().contentType(contentType).accept(contentType).config(RestAssured.config()
                .encoderConfig(encoderConfig().defaultContentCharset("UTF-8").appendDefaultContentCharsetToContentTypeIfUndefined(false)));
    }

    public API (String domain,String endPoint,String authorization){
        this.domain =domain;
        this.endPoint=endPoint;
        spec = rest().contentType(ContentType.JSON).accept(ContentType.JSON).header(AUTHORIZATION,authorization)
                .config(RestAssured.config()
                .encoderConfig(encoderConfig().defaultContentCharset("UTF-8").appendDefaultContentCharsetToContentTypeIfUndefined(false)));
    }

    public void setBasicAuth(String username,String password){
        spec.auth().preemptive().basic(username,password);
    }

    public void setHeader(String key,String value){
        spec.header(key,value);
    }

    public void setContentType(ContentType ct){
        spec.contentType(ct);
    }

    public void setParam(String key,String value){
        spec.param(key,value);
    }

    public Response getResponse(){return response;}

    public void submitRequest(Method method){
        String fullPath = domain+endPoint;
        fire(method,fullPath);
    }

    public void fire(Method method,String fullPath){
        LOGGER.info(fullPath.concat("[".concat(method.toString().concat("]"))));
        switch (method.toString()){
            case "POST":
                response = spec.post(fullPath);
                break;
            case "DELETE":
                response = spec.delete(fullPath);
                break;
            case "PUT":
                response = spec.put(fullPath);
                break;
            default:
                response = spec.get(fullPath);
                break;
        }
        if (failRetryCode.contains(response.getStatusCode())&& attempt==maxRetryTime){
//            Evidence.recordTextEvidence();
        }
    }

    public void validateResponseCode(){response.then().assertThat().statusCode(Matchers.isIn(successCode));}

    private String getAPIDetails(){
        return "Url: "+ domain+endPoint+"\n"
                +"Request :  \n"+payload+"\n"
                +"Response : \n"+response.asString()+"\n";
    }

    public int getStatusCode(){return response.getStatusCode();}

    public void ignoreSSL(boolean ignore){
        if(ignore){
            spec.relaxedHTTPSValidation();
        }
    }

    public void setPayload(String payload){
        this.payload=payload;
        spec.body(payload);
    }

    public void setResponse(Response response){
        this.response=response;
    }

    public <T> void validate(String path, Matcher<T> matcher){
        LOGGER.info("Validating path : "+path+"and matcher: "+matcher.toString());
        response.then().assertThat().body(path,matcher);
    }

    public void setFormParam(String... keyValues){
        if (keyValues.length % 2 !=0){
            LOGGER.debug("Key Value pair is not complete, hence form is not getting set");
        }else {
            for (int i=0;i<keyValues.length;i+=2){
                spec.formParam(keyValues[i],keyValues[i+1]);
            }
        }
    }

    public <T>List<T> getListAsObject(String path,Class<T> genericType){
        String contentType = response.getContentType();
        if(contentType.contains("json")){
            return RestAssuredCustomUtil.getListAsObjectJson(path,genericType,response.jsonPath());
        }else{
            return RestAssuredCustomUtil.getListAsObjectXml(path,genericType,response.xmlPath());
        }
    }
}
