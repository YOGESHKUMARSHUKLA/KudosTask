package starter.generic;

import org.apache.log4j.Logger;

import java.nio.charset.Charset;
import java.util.Map;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Optional;
import java.util.stream.Collectors;

public class InputTemplate implements InputTemplateInterface{
    private static final Logger LOGGER = Logger.getLogger(InputTemplate.class);
    private Path path;
    private String template;
    private String templateRef;
    private String stringPath;

    public InputTemplate(String path){
        this.path=Paths.get(path);
        read();
    }

    public InputTemplate(String initPath,String path){
        this.stringPath=initPath+path;
        readFromPath();
    }

    public void replaceVariable(String key,String value){
        LOGGER.debug("Replacing key: "+key+",with value: "+value);
        template=template.replaceAll("\\{\\{"+key+"\\}\\}",value);
    }

    public void replaceAllVariable(Map<String,String> dictionary){
        for (String key : dictionary.keySet()){
            String value = Optional.ofNullable(dictionary.get(key)).orElse("null");
            replaceVariable(key,value);
        }
    }

    public void replaceSubTree(String treeKey,String subTree){
        LOGGER.debug("Replacing treeKey: "+treeKey+",with value: "+subTree);
        template=template.replaceAll("<!-- "+treeKey+" -->",subTree);
    }

    @Override
    public void replaceAllSubTree(Map<String, String> dictionary) {

    }

    public void read() {
        LOGGER.info("Reading template from: "+path);
        InputStream inputStream;
        if(System.getProperty("os.name").startsWith("Windows")){
            inputStream =this.getClass().getClassLoader().getResourceAsStream(path.toString());
        }else{
            inputStream =this.getClass().getResourceAsStream(path.toString());
        }
        InputStreamReader is = new InputStreamReader(inputStream, Charset.forName("UTF-8"));
        BufferedReader buffer =new BufferedReader(is);
        templateRef=buffer.lines().collect(Collectors.joining("\n"));
        reset();
        LOGGER.debug("Read Template");
        LOGGER.debug(template);
    }

    public void readFromPath() {
        LOGGER.info("Reading template from: "+path);
        InputStream inputStream=null;
        try{
            inputStream =new FileInputStream(stringPath);
        }catch (Exception e){

        }
        InputStreamReader is = new InputStreamReader(inputStream, Charset.forName("UTF-8"));
        BufferedReader buffer =new BufferedReader(is);
        templateRef=buffer.lines().collect(Collectors.joining("\n"));
        reset();
        LOGGER.debug("Read Template");
        LOGGER.debug(template);
    }

    public void reset(){
        LOGGER.debug("Resetting template");
        template=templateRef;
    }

    public String getTemplate(){
        return template.replaceAll("\\{\\{[^\\}]*\\}\\}","").replaceAll("<!--[^>]*-->","");
    }
}
