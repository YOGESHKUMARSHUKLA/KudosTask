package starter.generic;

import java.util.Map;

public interface InputTemplateInterface {
    void replaceVariable(String key,String value);
    void replaceAllVariable(Map<String,String> dictionary);
    void replaceSubTree(String key,String subTree);
    void replaceAllSubTree(Map<String,String> dictionary);
    void read();
    String getTemplate();
}
