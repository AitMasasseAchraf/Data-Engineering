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

public class Reader_xls {
	/* le chemin de fichier sql*/
	private String path;
	/* le nom acronym de filiere*/
	private String acronym_none;
	
	public  Reader_xls(String path,String acronym_none) {
		this.path=path;
		this.acronym_none=acronym_none;
	}
	

   /* mutateur d'acronyme de filiere  */
	public String getAcronym_none() {
		return acronym_none;
	}
	
	
	
	public void setAcronym_none(String acronym_none) {
		this.acronym_none = acronym_none;
	}
	
	
	public String getPath() {
		return path;
	}
	
	public void setPath(String path) {
		this.path = path;
	}
	
	public HSSFSheet sheet_obj() throws IOException{
		/* j'utilise FileInputStream pour obtenir path file*/
		FileInputStream fis=new FileInputStream(new File(getPath()));  
		/* on implémente l’interface HSSFWorkbook pour les fichier excel de format xls*/
		
		HSSFWorkbook wb=new HSSFWorkbook(fis);
		/* on utilise HSSFSheet object pour lire et ecrire dans un fichier excel*/
		HSSFSheet sheet=wb.getSheetAt(0);
		return sheet;
				
	}
  public FormulaEvaluator evaluating()  throws IOException{
	  /* j'utilise FileInputStream pour obtenir path file*/
		FileInputStream fis=new FileInputStream(new File(getPath()));
		/* on implémente l’interface HSSFWorkbook pour les fichier excel de format xls*/
		HSSFWorkbook wb=new HSSFWorkbook(fis); 
		/* on utilise HSSFSheet object pour lire et ecrire dans un fichier excel*/
		 FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();
		 /* FormulaEvaluator nous permet d'evaluer la valeur de chaque cellule */
		 return formulaEvaluator;
	
}
  public void MOD_ELEM_ID( String path_file) {
	  try {
			 

          FileWriter fWriter = new FileWriter(
              path_file,true);
          /* le parameter true (boolean) permet d'ajouter un contenu sur le contenu existé sur le fichier sql*/

int nb_ann=0;
for(Row row: this.sheet_obj())     //iteration over row using for each loop  
{  int n_sm=0; int nb_mod=0;
for(Cell cell: row)    //iteration over cell using for each loop  
{  
this.evaluating().evaluateInCell(cell).getCellType();
// pour evaluer le type de chaque cellule //

if ( cell.getCellType()==Cell.CELL_TYPE_NUMERIC ) {
	 n_sm=(int)cell.getNumericCellValue();
	 
	 if(n_sm%2==1) {
		nb_ann++;
		String text1="---"+nb_ann+"ére Année\n";
		 String text2="insert into Mod_Elem_id VALUES ('HI"+this.acronym_none+String.valueOf(n_sm)+"04','ENH','AN' ,'','VET_ELEM_NOM','VET_ELEM_NOM','','23-JUL-2014')\n";
		 fWriter.write(text1);
		 fWriter.write(text2);
		
		 
	 }
	 String text="--S"+String.valueOf(n_sm)+"\n";
	 fWriter.write(text);

	String text3="insert into Mod_Elem_id VALUES ('HI"+this.acronym_none+String.valueOf(n_sm)+"004','ENH','SM0"+String.valueOf(n_sm)+"','S"+String.valueOf(n_sm)+"','semestre"+String.valueOf(n_sm)+"','semestre"+String.valueOf(n_sm)+"'','23-jul-2014')\n";
	 fWriter.write(text3);
}
if (cell.getCellType()==Cell.CELL_TYPE_STRING) {
	nb_mod++;
	String text4="insert into Mod_Elem_id VALUES ('HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(nb_mod)+"04','ENH','MOD"+nb_mod+"','S"+n_sm+"','"+cell.getStringCellValue()+"','"+cell.getStringCellValue()+"')\n";
	 fWriter.write(text4);
	String word = "/";
	 String sentence=cell.getStringCellValue();
	String temp[] = sentence.split(" ");
	int count=0;
	for (int i = 0; i < temp.length; i++) {
		if (word.equals(temp[i]))
		count++;
		}
	if(count!=0) {
		int nb_elem=0;
		for(int i=1 ; i<=count+1;i++) {
			String text5="insert into Mod_Elem_id VALUES ('HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(nb_mod)+String.valueOf(i)+"4','ENH','ELEM','S"+String.valueOf(n_sm)+"','"+temp[nb_elem]+"','"+temp[nb_elem]+"')\n\n\n";
			 fWriter.write(text5);
			nb_elem+=2; 
			 
		}
	}

} 
}
} fWriter.close();
System.out.println(
       "MOD_ELEM_ID is created successfully with the content.\n");
} catch (IOException e) {
	 
  // Print the exception
  System.out.print(e.getMessage());
}
  }
  public void List_Mod_Elm_Id(String path_file) {
	  try {
		  
	  
	  int nb_semestre=1;
	  FileWriter fWriter = new FileWriter(
              path_file,true);
		for(Row row: this.sheet_obj())     //iteration over row using for each loop  
		{  int s=0; int n_sm=0;
		for(Cell cell: row)    //iteration over cell using for each loop  
		{  
		this.evaluating().evaluateInCell(cell).getCellType();
		if ( cell.getCellType()==Cell.CELL_TYPE_NUMERIC ) {
			 n_sm=(int)cell.getNumericCellValue();
			 if(n_sm%2==1) {
				 String text1="insert into List_Mod_Elm_Id VALUES('HI"+this.acronym_none+String.valueOf(nb_semestre)+"4' ,'VET "+
			 String.valueOf(nb_semestre)+"année ID','VET "+
					 String.valueOf(nb_semestre)+"année ID')\n";
				
				 String text2="insert into List_Mod_Elm_Id VALUES('HI"+this.acronym_none+String.valueOf(nb_semestre)+"04','VET_ELEM_NOM','VET_ELEM_NOM')\n";
				 fWriter.write(text1);
				 fWriter.write(text2);
				 nb_semestre++;
				 
			 }
	
			String text="insert into List_Mod_Elm_Id VALUES('HI"+this.acronym_none+String.valueOf(n_sm)+"004','semestre"+String.valueOf(n_sm)+"','semestre"+String.valueOf(n_sm)+"')\n";
			 fWriter.write(text);
			
	} 
		if (cell.getCellType()==Cell.CELL_TYPE_STRING) { 
			String text="insert into List_Mod_Elm_Id VALUES('HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(s)+"04,'"+cell.getStringCellValue()+"','"+cell.getStringCellValue()+"')\n\n\n";
			 fWriter.write(text);
  }
		}
		s++;
		
		
	}
		fWriter.close();
		System.out.println(
			       "LIST_MOD_ELEM_ID is created successfully with the content.\n");
		}  catch (IOException e) {
			 
			  // Print the exception
			  System.out.print(e.getMessage());
			}}
  
  public void Trait_Notes_Id (String path_file) {
	  try {
		  
	  
	  int nb_semestre=1;
	  FileWriter fWriter = new FileWriter(
              path_file,true);
		for(Row row: this.sheet_obj())     //iteration over row using for each loop  
		{  int s=0; int n_sm=0;
		System.out.println("\n");
		for(Cell cell: row)    //iteration over cell using for each loop  
		{  
		this.evaluating().evaluateInCell(cell).getCellType(); 
		
			if ( cell.getCellType()==Cell.CELL_TYPE_NUMERIC ) {
				 n_sm=(int)cell.getNumericCellValue();
				 if(n_sm%2==1) {
					 String text1="insert into Trait_Notes_Id values ('HI"+this.acronym_none+String.valueOf(nb_semestre)+"04','HTN','T')\n";
					 fWriter.write(text1);
					 
					 nb_semestre++;
					 
				 }
				String text="insert into Trait_Notes_Id values ('HI"+this.acronym_none+String.valueOf(n_sm)+"004','HTN','T')\n";
				 fWriter.write(text);
		} 
			if (cell.getCellType()==Cell.CELL_TYPE_STRING) { 
				String text1="insert into Trait_Notes_Id values ('HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(s)+"04','HTN','T')\n";
				 fWriter.write(text1);
				
//				
				String word = "/";
				 String sentence=cell.getStringCellValue();
				String temp[] = sentence.split(" ");
				int count=0;
				for (int i = 0; i < temp.length; i++) {
					if (word.equals(temp[i]))
					count++;
					}
				if(count!=0) {
					for(int i=1 ; i<=count+1;i++) {
						String text="insert into Trait_Notes_Id values ('HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(s)+String.valueOf(i)+"4','HTN','T')\n\n\n";
						 fWriter.write(text);
					}
				}
		}
			s++;
			
			
		}}fWriter.close();
		System.out.println(
			       "TRait_ELEM_ID is created successfully with the content.\n");
	  
  } catch (IOException e) {
		 
	  // Print the exception
	  System.out.print(e.getMessage());
	}}
  
  
  
  public void Inscr_Pedag_Id(String path_file) {
	  try {
		  
	  
	  int nb_semestre=1;
	  FileWriter fWriter = new FileWriter(
              path_file,true);
		for(Row row: this.sheet_obj())     //iteration over row using for each loop  
		{  int s=0; int n_sm=0;
		for(Cell cell: row)    //iteration over cell using for each loop  
		{  
		this.evaluating().evaluateInCell(cell).getCellType();
		if ( cell.getCellType()==Cell.CELL_TYPE_NUMERIC ) {
			 n_sm=(int)cell.getNumericCellValue();
			 if(n_sm%2==1) {
				
				 String text2="insert into Inscr_Pedag_Id values ('ENH','HI"+this.acronym_none+String.valueOf(nb_semestre)+"04')\n";
				 fWriter.write(text2);
				 nb_semestre++;
				 
			 }
	
			String text="insert into Inscr_Pedag_Id values ('ENH','HI"+this.acronym_none+String.valueOf(n_sm)+"004')\n";
			 fWriter.write(text);
			
	} 
		if (cell.getCellType()==Cell.CELL_TYPE_STRING) { 
			String text="insert into Inscr_Pedag_Id VALUES('ENH','HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(s)+"04)\n";
			 fWriter.write(text);

				String word = "/";
				 String sentence=cell.getStringCellValue();
				String temp[] = sentence.split(" ");
				int count=0;
				for (int i = 0; i < temp.length; i++) {
					if (word.equals(temp[i]))
					count++;
					}
				if(count!=0) {
					for(int i=1 ; i<=count+1;i++) {
						String text3="insert into Inscr_Pedag_Id values ('ENH','HI"+this.acronym_none+String.valueOf(n_sm)+String.valueOf(s)+String.valueOf(i)+"4')\n\n\n";
						 fWriter.write(text3);
					}
				}
  }s++;
		}
	
		
		
	}
		fWriter.close();
		System.out.println(
			       "Inscr_Pedac_Id is created successfully with the content.");
		}  catch (IOException e) {
			 
			  // Print the exception
			  System.out.print(e.getMessage());
			}}
	

}
