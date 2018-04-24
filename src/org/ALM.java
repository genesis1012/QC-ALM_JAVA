package org;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import HP.ExcelApiTest;

import com.qc.ClassFactory;
import com.qc.IAttachment;
import com.qc.IAttachmentFactory;
import com.qc.IBaseFactory;
import com.qc.IBug;
import com.qc.IBugFactory;
import com.qc.ICycle;
import com.qc.ICycleFactory;
import com.qc.IExtendedStorage;
import com.qc.ILink;
import com.qc.ILinkFactory;
import com.qc.ILinkable;
import com.qc.IList;
import com.qc.IRelease;
import com.qc.IReleaseFactory;
import com.qc.IRun;
import com.qc.IRunFactory;
import com.qc.ITDConnection4;
import com.qc.ITDFilter;
import com.qc.ITSTest;
import com.qc.ITestSet;
import com.qc.ITestSetFactory;

import com4j.Com4jObject;
import com4j.Variant;

public class ALM {
	
	Date date = new Date();
	DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
	DateFormat dateFormatH = new SimpleDateFormat("dd_MM_yyyy_HH_mm_ss");
	DateFormat dateFormatR = new SimpleDateFormat("yyyy HH:mm:ss");
	String fecha = dateFormatH.format(date);
	String Rum = dateFormatR.format(date);
	Utils.XmlUtils Parameter = new Utils.XmlUtils();
	ExcelApiTest exc = null;
	private String Dir = System.getProperty("user.dir");
	
	/********************Procedimiento de conexion a ALM******************/
	private ITDConnection4 QCConecctio(){
		ITDConnection4 QCConnection = null;
		try{
			QCConnection = ClassFactory.createTDConnection();
			QCConnection.initConnectionEx(Parameter.leerNodo("CadenaConexion"));//Link para acceso a ALM/QC 12
			QCConnection.login(Parameter.leerNodo("usuarioALM"),Parameter.leerNodo("contrasenaALM"));
			QCConnection.connect(Parameter.leerNodo("dominio"), Parameter.leerNodo("subdominio"));
			
			exc = new ExcelApiTest(Dir + Parameter.leerNodo("pathDefects"));
			
		}catch(Exception e){
			System.out.println("Exceptions occured: "+e.getMessage());
		}
		return QCConnection;
	}
	
	/*******************Crear defecto*******************/
	public String createDefect(String Val_id, String fallo)  {
	    IBugFactory  bugFactory = (IBugFactory) QCConecctio().bugFactory().queryInterface(IBugFactory.class);//  connection.bugFactory().queryInterface(IBugFactory.class);
	    IBug bug = (bugFactory.addItem(new Variant(Variant.Type.VT_NULL))).queryInterface(IBug.class);
	    
	    bug.summary(exc.getCellData("summary", Val_id));//descripcion general del BUG
	    bug.status("New"); //Estado del bug
	    bug.assignedTo(exc.getCellData("assignedTo", Val_id)); //Asignado a
	    bug.detectedBy(exc.getCellData("detectedBy", Val_id)); //Detectado por
	    bug.field("BG_DETECTION_DATE", dateFormat.format(date)); //Fecha de deteccion del bug
	    bug.field("BG_DETECTED_IN_REL", Parameter.leerNodo("Releases")); //ID del Releases SOLUCION
	    bug.field("BG_DETECTED_IN_RCYC", Parameter.leerNodo("Sprint")); //ID del Sprint dentro del Releases SOLUCION
	    bug.field("BG_USER_TEMPLATE_06", Parameter.leerNodo("Testing"));//proveedor de testing
	    bug.field("BG_USER_TEMPLATE_08", exc.getCellData("Responsable", Val_id)); //proveedor responsable
	    bug.field("BG_USER_TEMPLATE_12", Parameter.leerNodo("Aplicativo")); //aplicativo
	    bug.field("BG_USER_TEMPLATE_07", "Desarrollo"); //Naturalesa
	    bug.field("BG_USER_TEMPLATE_09", "Error"); //tipo de defecto
	    bug.field("BG_USER_TEMPLATE_01", exc.getCellData("Impacto", Val_id)); //Impacto al negocio
	    bug.field("BG_USER_TEMPLATE_02", exc.getCellData("Probabilidad", Val_id)); //probabilidad de falla
	    bug.field("BG_DESCRIPTION", "Defecto automatico - "+ exc.getCellData("BG_DESCRIPTION", Val_id) + ": " + fallo); //descripn detallada del error
	    bug.field("BG_USER_TEMPLATE_03", "Y"); //Defecto automatico
	    
	     
	    bug.post(); 
	    System.out.println("Defecto creado");
	    Update_Bug(bug.id().toString(), "Open", Val_id, " ");
	    return bug.id().toString();
	}
	
	/*******************Ejecutar Test en TestLab******************************/
	public void rum(String TestSet, String Test, String strStatus){
		IRunFactory runfactory = Fin_Test_TestSet_Rum(TestSet, Test).runFactory().queryInterface(IRunFactory.class);
		IRun run= runfactory.addItem(Rum).queryInterface(IRun.class);
		run.status(strStatus);
		run.post();  
		System.out.println("CP Ejecutado");
	}
	
	/*****************Buscar Defecto**********************/
	private IBug Buscar_Bug(String id){
	    
	    IBugFactory bugfactory = QCConecctio().bugFactory().queryInterface(IBugFactory.class);
	    IBug bug = bugfactory.item(id).queryInterface(IBug.class);
	    return bug;
	}
	
	/***********************Actualizar Defecto*********************/
	public void Update_Bug(String id, String status, String Val_id, String Comen){
		
	    IBugFactory bugfactory = QCConecctio().bugFactory().queryInterface(IBugFactory.class);
	    IBug bug = bugfactory.item(id).queryInterface(IBug.class);
	    bug.status(status);
	    
	    String comen = "<br/>" + exc.getCellData("detectedBy", Val_id) + ", " + dateFormat.format(date) + ": " + Comen;
	    bug.summary(bug.summary());//descripcion general del BUG
	    bug.assignedTo(bug.assignedTo()); //Asignado a
	    bug.detectedBy(bug.detectedBy()); //Detectado por
	    bug.field("BG_DETECTION_DATE", bug.field("BG_DETECTION_DATE")); //Fecha de deteccion del bug
	    bug.field("BG_USER_TEMPLATE_06", bug.field("BG_USER_TEMPLATE_06"));//proveedor de testing
	    bug.field("BG_USER_TEMPLATE_08", exc.getCellData("Responsable", Val_id)); //proveedor responsable
	    bug.field("BG_USER_TEMPLATE_12", bug.field("BG_USER_TEMPLATE_12")); //aplicativo
	    bug.field("BG_USER_TEMPLATE_07", bug.field("BG_USER_TEMPLATE_07")); //Naturalesa
	    bug.field("BG_USER_TEMPLATE_09", bug.field("BG_USER_TEMPLATE_09")); //tipo de defecto
	    bug.field("BG_USER_TEMPLATE_01", bug.field("BG_USER_TEMPLATE_01")); //Impacto al negocio
	    bug.field("BG_USER_TEMPLATE_02", bug.field("BG_USER_TEMPLATE_02")); //probabilidad de falla
	    //bug.field("BG_USER_TEMPLATE_02", bug.field("BG_USER_TEMPLATE_02")); //probabilidad de falla
	    bug.field("BG_DEV_COMMENTS", bug.field("BG_DEV_COMMENTS") + comen); //Comentario del Defecto
	    
	    bug.post();
	    System.out.println("Defecto Actualizado");
	}
	
	/******************Buscar Defecto*****************/
	public String Bug_Status(String id){
		
		IBugFactory bugfactory = QCConecctio().bugFactory().queryInterface(IBugFactory.class);
		ITDFilter fil = bugfactory.filter().queryInterface(ITDFilter.class);
		fil.filter("BG_BUG_ID",id); //any filter value
		IList buglist = fil.newList();      
		Iterator itr = buglist.iterator();
		IBug bug=null;
		while(itr.hasNext()){
		   Com4jObject comobj = (Com4jObject)itr.next();
		   bug = comobj.queryInterface(IBug.class);
		}	
		return bug.status().toString();
	}
	
	/********************Buscar Sprint en un Releases*********************/
	private String Find_Sprint_Releases(String SprintName, String id_reles){
		IReleaseFactory rel = QCConecctio().releaseFactory().queryInterface(IReleaseFactory.class);
		ITDFilter fil = rel.filter().queryInterface(ITDFilter.class);
		fil.filter("REL_ID", id_reles);
		IList buglist = fil.newList();      
		Iterator itr = buglist.iterator();
		IRelease reles = null;
		while(itr.hasNext()){
		   Com4jObject comobj = (Com4jObject)itr.next();
		   reles = comobj.queryInterface(IRelease.class);
		}
		
		ICycleFactory obj2 = reles.cycleFactory().queryInterface(ICycleFactory.class);
		IList tstestlist = obj2.newList("");
		ICycle Sprint = null;
		for(Com4jObject obj3:tstestlist)
		{
			ICycle tstest = obj3.queryInterface(ICycle.class);
			if(tstest.name().contains(SprintName)){	
				Sprint = tstest;
				break;
			}
		}	
		return Sprint.id().toString();
	}
	
	/********************Buscar Releases***********************/
	private String Find_Releases(String id_reles){
		IReleaseFactory rel = QCConecctio().releaseFactory().queryInterface(IReleaseFactory.class);
		ITDFilter fil = rel.filter().queryInterface(ITDFilter.class);
		fil.filter("REL_ID", id_reles);
		IList buglist = fil.newList();      
		Iterator itr = buglist.iterator();
		IRelease reles = null;
		while(itr.hasNext()){
		   Com4jObject comobj = (Com4jObject)itr.next();
		   reles = comobj.queryInterface(IRelease.class);
		}	
		return reles.id().toString();
	}

	/************************Linked TestCase a Bud*****************************/
	public void linked(String strTestSetName, String strTestCasename, String bug_id){		
		ITSTest b2Test = Fin_Test_TestSet_Rum(strTestSetName, strTestCasename);
		IBug bug = Buscar_Bug(bug_id);
		
		ILinkFactory linkF = bug.queryInterface(ILinkFactory.class);
		ILink alink = QCConecctio().bugFactory().queryInterface(ILink.class);
		ILinkable iLinkable = b2Test.queryInterface(ILinkable.class);
		linkF = iLinkable.bugLinkFactory().queryInterface(ILinkFactory.class);
		alink = linkF.addItem(bug.id()).queryInterface(ILink.class);
		alink.linkType("Related");
		alink.post();
		System.out.println("Defecto Linkeado al CP");
	}	
	
	/*******************Buscar test en Test set TestLab***************************/
	private ITSTest Fin_Test_TestSet_Rum(String ID_Testset, String strTestCasename){
		boolean existTest = false;
		ITSTest objTestcase =null;
		ITestSetFactory sTestSetFactory = (QCConecctio().testSetFactory()).queryInterface(ITestSetFactory.class);    
		ITestSet sTestSet = (sTestSetFactory.item(ID_Testset)).queryInterface(ITestSet.class); 
		existTest = true;

			if(existTest){
			IBaseFactory obj2 = sTestSet.tsTestFactory().queryInterface(IBaseFactory.class);
			IList tstestlist = obj2.newList("");
			boolean testcaseExist = false;
			ITSTest testCase = null;
			for(Com4jObject obj3:tstestlist)
			{
				ITSTest tstest = obj3.queryInterface(ITSTest.class);
				System.out.println(tstest.name());
				System.out.println(tstest.id());
				if(tstest.name().contains(strTestCasename)){
				//if(tstest.id().equals(strTestCasename)){
					testcaseExist = true;
					testCase = tstest;
					break;
				}
			}
			if(testcaseExist){
				System.out.println("Testcase \""+ strTestCasename +"\" Found under the Test Set Name \"" + ID_Testset +"\"");
				objTestcase = testCase;
				
			}else{
				objTestcase = null;
				System.out.println("Testcase \""+ strTestCasename +"\" not Found under the Test Set Name \"" + ID_Testset + "\"");
			}
		}else{
			System.out.println("Test Set Name \"" + ID_Testset + "\"");
		}
		return objTestcase;
	}
	
	/***********************Adjuntar evidencias a un Defecto******************************/
	public void adjuntar(String bug_id, String fileName, String folderName){//(IBug run){
		try{
			String filename = fileName+"_"+fecha+".png";
            IAttachmentFactory attachfac = Buscar_Bug(bug_id).attachments().queryInterface(IAttachmentFactory.class);
		     IAttachment attach = attachfac.addItem(filename).queryInterface(IAttachment.class);
		     IExtendedStorage extAttach = attach.attachmentStorage().queryInterface(IExtendedStorage.class);
		     extAttach.clientPath(folderName); 
		     extAttach.save(filename, true);
		     attach.description("File");
		     //attach.type(1);
		     attach.post();
		     attach.refresh();
		     System.out.println("Evidencia Adjunta");
		    }catch(Exception e) {
		     System.out.println("QC Exceptione : "+e.getMessage());
		    }	
	}
	
	/*********************Tomar evidencas]
	 * @throws InterruptedException *******************************/
	public void takeScreenShotTest(WebDriver driver, String imageName, String folderName, WebElement element) throws InterruptedException {
		
		highlightElement(driver, element);
		
		String filename = imageName+"_"+fecha+".png";
		File directory = new File(folderName);
	      try {
	         if (directory.isDirectory()) {
	            //Toma la captura de imagen
	            File imagen = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
	            //Mueve el archivo a la carga especificada con el respectivo nombre
	            FileUtils.copyFile(imagen, new File(directory.getAbsolutePath()   + "\\" + filename));
	            System.out.println("Evidencia tomada");
	         } else {
	            //Se lanza la excepcion cuando no encuentre el directorio
	            throw new IOException("ERROR : La ruta especificada no es un directorio!");
	         }
	      } catch (IOException e) {
	         //Impresion de Excepciones
	         e.printStackTrace();
	      }
	   }
	
	/*********************Iluminar Objeto**************************/
	public void highlightElement(WebDriver driver, WebElement element) throws InterruptedException {
		JavascriptExecutor js=(JavascriptExecutor)driver;
		for (int i = 0; i < 3; i++) {
			js.executeScript("arguments[0].setAttribute('style', arguments[1]);",
            element, "color: red; border: 5px solid red;");
            Thread.sleep(2000);
            //js.executeScript("arguments[0].setAttribute('style', arguments[1]);",element, "");
        }
	}

}
