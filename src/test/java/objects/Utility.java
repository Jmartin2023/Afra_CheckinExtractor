package objects;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

public class Utility {
	
	public static void main(String[] args) throws SAXException, IOException, ParserConfigurationException {
		
	}
	
	public String toPascalCase(String text) {
        if (text == null || text.isEmpty()) {
            return text;
        }

        StringBuilder converted = new StringBuilder();

        boolean convertNext = true;
        for (char ch : text.toCharArray()) {
            if (Character.isSpaceChar(ch)) {
                convertNext = true;
            } else if (convertNext) {
                ch = Character.toTitleCase(ch);
                convertNext = false;
            } else {
                ch = Character.toLowerCase(ch);
            }
            converted.append(ch);
        }

        return converted.toString();
    }
	
	public HashMap<String, String> getConfig(String path, String[] params) throws SAXException, IOException, ParserConfigurationException {
		
		File xmlFile = new File(path);
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = factory.newDocumentBuilder();
		Document doc = builder.parse(xmlFile);
		
		Element credentialElement = (Element) doc.getElementsByTagName("config").item(0);
	    
		HashMap<String, String> configs = new HashMap<String, String>(params.length);
		for(String param : params) {
			String text = credentialElement.getElementsByTagName(param).item(0).getTextContent();
			configs.put(param, text);
		}
		
		return configs;		
	}

}
