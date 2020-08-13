package xmlParseExcel;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class exceldata {
	
	public static final String EXCELFILELOCATION= "C:\\Software\\learning workspac\\Evidence\\ExcelToXml\\Test.xlsx";
	private static FileInputStream fis;
	 private static XSSFWorkbook workbook;
	 private static XSSFSheet sheet;
	 private static XSSFRow row;
	 
	 public static void loadExcel() throws Exception{
		 	
		 File file = new File(EXCELFILELOCATION);
		 fis=new FileInputStream(file);
		 workbook = new XSSFWorkbook(fis);
		 sheet=workbook.getSheet("Test");
		 fis.close();
		 
	 }
	public static List<Map<String, String>> readalldata() throws Exception{
		if(sheet==null) {
			 loadExcel();
		 }
		 List<Map<String, String>> testdataallrows=null;
		Map<String, String> testdata = null;
		
		
		int row=sheet.getLastRowNum(); //find number of row
		int cell=sheet.getRow(0).getLastCellNum(); //find number of colmn
		
		List list = new ArrayList();
		for(int i=0;i<cell;i++ ) {
		    Row r=sheet.getRow(0);
		   Cell c= r.getCell(i);
		  String rowheader= c.getStringCellValue(); //readrow value
		  list.add(rowheader);
		}
		testdataallrows = new ArrayList<Map<String, String>>(); 
		
		for(int j=1;j<=row;j++) {
			Row r = sheet.getRow(j);
			testdata =new TreeMap<String, String>(String.CASE_INSENSITIVE_ORDER); //keep insertion order
			for(int k=0;k<cell;k++)
			{
				Cell c=r.getCell(k); //read rest of the row
				String colvalue= c.getStringCellValue();
				testdata.put((String) list.get(k), colvalue); 
			}
			testdataallrows.add(testdata);
		}
		return testdataallrows;
	 }
	
	
	public static String getValue(String key) throws Exception{
		Map<String, String> myval = readalldata().get(4);
		 String retrive = myval.get(key);
		return retrive;
		 
	 }
	
	public void count() throws Exception {
		if(sheet==null) {
			 loadExcel();
		 }
		int rownum=sheet.getLastRowNum()+1;
		System.out.println("totol row" +rownum);
		Row row=sheet.getRow(0);
		int colnum=row.getLastCellNum();
		System.out.println("total col" +colnum);
	}
		
	
	
	public void readxml() throws Exception {
		
		File file = new File("C:\\Software\\learning workspac\\Evidence\\ExcelToXml\\Test.xml");
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder(); 
		Document doc = db.parse(file); 
		doc.getDocumentElement().normalize(); 
		System.out.println("Root element: " + doc.getDocumentElement().getNodeName());  
		NodeList nodeList = doc.getElementsByTagName("note");  
		for (int itr = 0; itr < nodeList.getLength(); itr++) 
		{  
			Node node = nodeList.item(itr);  
			System.out.println("\nNode Name :" + node.getNodeName());  
			if (node.getNodeType() == Node.ELEMENT_NODE)   
			{  
			Element eElement = (Element) node;  
			System.out.println("Description : "+ eElement.getElementsByTagName("Description").item(0).getTextContent());  
			System.out.println("Value: "+ eElement.getElementsByTagName("Value").item(0).getTextContent()); 
			System.out.println("Title: "+ eElement.getElementsByTagName("Title").item(0).getTextContent());  
			System.out.println("body: "+ eElement.getElementsByTagName("body").item(0).getTextContent());  
			System.out.println("base: "+ eElement.getElementsByTagName("base").item(0).getTextContent());  
			}  
			}  
	}
	
	public  void CreateXmlFile() throws Exception {
		
		 DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		 DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		 Document doc = dBuilder.newDocument();
		 Element rootElement = doc.createElement("note");
        doc.appendChild(rootElement);
        
       
        {
        
       	 
       	     Element d = doc.createElement("Description");
		     d.appendChild(doc.createTextNode("{" + getValue("Description") + "}"));
	         rootElement.appendChild(d);
        
        
	         
	         Element v = doc.createElement("Value");
		     v.appendChild(doc.createTextNode("{" + getValue("Value") + "}"));
	         rootElement.appendChild(v);
        
	         Element t = doc.createElement("Title");
		     t.appendChild(doc.createTextNode("{" + getValue("Title") + "}"));
	         rootElement.appendChild(t);
	         
	         Element b = doc.createElement("body");
		     b.appendChild(doc.createTextNode("{" + getValue("Browser") + "}"));
	         rootElement.appendChild(b);
	         
	         Element p = doc.createElement("base");
		     p.appendChild(doc.createTextNode("{" + getValue("Platform") + "}"));
	         rootElement.appendChild(p);
        
        }
        
               
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        DOMSource source = new DOMSource(doc);
        StreamResult result = new StreamResult(new File("C:\\Software\\learning workspac\\Evidence\\ExcelToXml\\testingout.xml"));
        transformer.transform(source, result);
        
        
        StreamResult consoleResult = new StreamResult(System.out);
        transformer.transform(source, consoleResult);
	} 


}
