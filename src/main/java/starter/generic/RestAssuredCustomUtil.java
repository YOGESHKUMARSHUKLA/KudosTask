package starter.generic;

import io.restassured.path.json.JsonPath;
import io.restassured.path.xml.XmlPath;
import org.apache.log4j.Logger;

import java.util.ArrayList;
import java.util.List;

public class RestAssuredCustomUtil {

    private static final Logger LOGGER = Logger.getLogger(RestAssuredCustomUtil.class);
    public static <T> List<T> getListAsObjectXml(String path, Class<T> genericType, XmlPath xmlpath){
        List<T> resultList;
        try{
            resultList=xmlpath.getList(path,genericType);
        }catch(
            ClassCastException e){
            LOGGER.error(e.getMessage());
            resultList =new ArrayList<>();
            T singleElement = xmlpath.getObject(path,genericType);
            resultList.add(singleElement);
        }
        return resultList;
    }

    public static <T> List<T> getListAsObjectJson(String path, Class<T> genericType, JsonPath jsonPath){
        List<T> resultList;
        try{
            resultList=jsonPath.getList(path,genericType);
        }catch(
                ClassCastException e){
            LOGGER.info(e.getMessage());
            resultList =new ArrayList<>();
            T singleElement = jsonPath.getObject(path,genericType);
            resultList.add(singleElement);
        }
        return resultList;
    }
}
