package pl_java;
import org.apache.log4j.Logger;
import java.io.File;  
import java.io.FileInputStream;  
import java.io.IOException;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;
import java.io.FileWriter;
import java.io.IOException;


public class main {
	// on crée une instance de la classe Logger pour créer un Logger avec le nom de la classe de notre subsysteme//  
	private static Logger LOGGER = Logger.getLogger(Reader_xls.class);
	public static void main(String args[]) throws IOException  
	{  
// logger.debug nous permet de lancer un message durant chaque lancement des fonctions entre try et catch//
		LOGGER.debug("Creating Sql Tables");
		try {
		
			Reader_xls obj=new Reader_xls("C:\\\\Users\\\\aitma\\\\Desktop\\\\Book1_2007.xls","GEE");
			/*LOGGER.debug(obj) nous permet de prendre  les informations qui peuvent être nécessaires au diagnostic des problèmes et 
			au dépannage ou lors de l’exécution de l’application dans l’environnement 
			de test afin de s’assurer que tout fonctionne correctement*/
			
			
			
			 LOGGER.debug(obj);
			 
			HSSFSheet sheet=obj.sheet_obj();
			
			FormulaEvaluator formulaEvaluator=obj.evaluating();
			String path="C:\\Users\\aitma\\Desktop\\example_file.txt";
			obj.MOD_ELEM_ID(path);
			obj.List_Mod_Elm_Id(path);
			obj.Trait_Notes_Id(path);
			obj.Inscr_Pedag_Id(path);
			
			
		
		}catch(Exception e) {
			
			//LOGGER>error nous permet de lancer le message erreur
			 LOGGER.error(e.getMessage(), e);
		}
		

		 }
}
