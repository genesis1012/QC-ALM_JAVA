package Inicial;

import java.io.IOException;

import org.junit.Assert;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import HP.ExcelApiTest;

import org.ALM;

public class Validar {
	
	ExcelApiTest exc;
	ALM ac = new ALM();
	Utils.XmlUtils param = new Utils.XmlUtils();
	public WebDriver driver=null;
	public WebElement Element=null;
	
	private String pathDefects;
	private String TestCase;
	private String TestSet;
	private String pathEvidencias;
	private String Dir = System.getProperty("user.dir");
	private String Comen_Pass, Comen_Fail;
	
	/*******************************************************************************************
	 * Creador: Heriberto Genes Garca
	 * Fecha: 05/03/2018
	 * Proveedor: TCS
	 * “El conocimiento no es una vasija que se llena, sino un fuego que se enciende”. Plutarco
	 * @throws Exception 
	 * ******************************************************************************************/
	
	public void ValidarTest(String esperado, String resultado, String ID_HU, WebDriver Driver, WebElement element) throws Exception{
		pathDefects = param.leerNodo("pathDefects");
		TestCase = param.leerNodo("TestCase");
		TestSet = param.leerNodo("TestSet");
		pathEvidencias = param.leerNodo("pathEvidencias");
		exc = new ExcelApiTest(Dir + pathDefects);
		Comen_Pass = exc.getCellData("Comen_Pass", ID_HU);
		Comen_Fail = exc.getCellData("Comen_Fail", ID_HU);
		
		try {
			driver = Driver;
			Element = element;	
			Assert.assertEquals( esperado, resultado);
			int id = Integer.valueOf(exc.getCellData("BUG_ID", ID_HU));
			if(id > 0){
				Bug_ALM("Si", ID_HU, "Solucionado");
			}
		} catch (AssertionError  e) {
			try{
				Bug_ALM("No", ID_HU, e.toString());
				Assert.assertFalse(true);
			}catch(Exception ex){
				ex.printStackTrace();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void Bug_ALM(String solucion, String HU, String Fallo){
		try {
		exc = new ExcelApiTest(Dir + pathDefects);
			String Bug_id = exc.getCellData("BUG_ID", HU); 
			if(Integer.valueOf(Bug_id) == 0){
				String id = ac.createDefect(HU, Fallo);
				exc.writeXLSXFile(Dir + pathDefects, HU, "BUG_ID", id);
				ac.rum(TestSet, TestCase, "Failed");
				ac.linked(TestSet, TestCase, id);
				ac.takeScreenShotTest(driver,HU,Dir + pathEvidencias, Element);
				ac.adjuntar(id,HU,Dir + pathEvidencias);
				System.out.println("el Val_ID:" + HU + ", se le asiga BUG_ID:" + id);
			}else if(Integer.valueOf(Bug_id) > 0){
				switch (ac.Bug_Status(Bug_id)) {
		            case "Fixed": 
		            	switch(solucion){
			            	case "Si":
			            		ac.Update_Bug(Bug_id, "Closed", HU, Comen_Pass);
			            		ac.rum(TestSet, TestCase, "Passed");
			            		break;
			            	case "No":
			            		ac.Update_Bug(Bug_id, "Reopen", HU, Comen_Fail);
			            		ac.rum(TestSet, TestCase, "Failed");
			            		ac.takeScreenShotTest(driver, HU,Dir + pathEvidencias, Element);
			            		ac.adjuntar(Bug_id,HU,Dir + pathEvidencias);
			        			break;
		            	}
		                break;
		            case "Closed":  
		            	switch(solucion){
		            	case "Si":
		            		//ac.Update_Bug(Bug_id, "Closed");
		            		break;
		            	case "No":
		            		ac.Update_Bug(Bug_id, "Reopen", HU, Comen_Fail);
		            		ac.rum(TestSet, TestCase, "Failed");
		            		ac.takeScreenShotTest(driver, HU, Dir + pathEvidencias, Element);
		            		ac.adjuntar(Bug_id,HU,Dir + pathEvidencias);
		        			break;
		            	}
		                break;
		            case "Rejected":  
		            	switch(solucion){
		            	case "Si":
		            		ac.Update_Bug(Bug_id, "Closed", HU, Comen_Pass);
		            		ac.rum(TestSet, TestCase, "Passed");
		            		break;
		            	case "No":
		            		ac.Update_Bug(Bug_id, "Rejected", HU, Comen_Fail);
		            		ac.rum(TestSet, TestCase, "Failed");
		            		ac.takeScreenShotTest(driver, HU,Dir + pathEvidencias, Element);
		            		ac.adjuntar(Bug_id,HU,Dir + pathEvidencias);
		        			break;
		            	}
		                break;
		            /*case "Reopen":  
		            	switch(solucion){
		            	case "Si":
		            		ac.Update_Bug(Bug_id, "Closed", HU);
		            		ac.rum(TestSet, TestCase, "Passed");
		            		break;
		            	case "No":
		            		/*ac.Update_Bug(Bug_id, "Rejected");
		            		ac.takeScreenShotTest(driver, "titulo_"+HU, param.leerNodo("pathEvidencias"));
		            		ac.adjuntar(Bug_id,"titulo_"+HU, param.leerNodo("pathEvidencias"));
		        			break;
		            	}
		                break;*/
				}
				
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
