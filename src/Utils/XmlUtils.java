package Utils;

import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;

public class XmlUtils {
	
	public String leerNodo(String tagName) {
		try {
			
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			dbFactory.setValidating(true);
			dbFactory.setExpandEntityReferences(false);
			
			Document doc = dbFactory.newDocumentBuilder().parse(System.getProperty("user.dir") + "\\parameters.xml");						
			
			String elemento = doc.getElementsByTagName(tagName).item(0).getTextContent();
			
			return elemento.trim().replace("\n", "");
						
		}catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
			return "";
		}
	}

}
