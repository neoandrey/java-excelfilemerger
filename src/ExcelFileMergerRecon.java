/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.text.SimpleDateFormat;


public class ExcelFileMergerRecon{

        String sourceDir;
        String TargetFileName;
		File TargetFile;
        File[] excelFiles;
        String oldFileName;
        String newFileName;
        File   excelSourceFolder;
        File   newFileFolder;
        String fileNamePrefix;
        Set<File>  fileSet = new HashSet<File>();
		//Set<File>  excelFiles = new HashSet<File>();
        Set<String>  fileNameSet = new HashSet<String>();
        Map<Integer, String[][]>  excelFileData = new HashMap<Integer, String[][]>();
        int dateIdentifier = 0;
        int renameCounter =0;
        File outputFile =  null;
        FileWriter fileWriter =  null;
        PrintWriter printWriter =  null;
        XSSFSheet mySheet;
		SimpleDateFormat df  = null;
		XSSFWorkbook mergedWorkbook = new XSSFWorkbook();
			int globalRow = 0;
	   DataFormatter formatter = new DataFormatter();

		public ExcelFileMergerRecon(){



		}

		public ExcelFileMergerRecon(String source, String target){
                 df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				 setSourceDir(source);
				 setTargetFileName(target);
				 initOutputFile();
				 getExcelFiles();
				 processExcelFiles();
				 writeMergedFile();
		}



	public void setSourceDir(String source){

	      this.sourceDir = source;

	}
	public void setTargetFile(File Target){

              this.TargetFile = Target;

	}
	public void setTargetFileName(String TargetName){

              this.TargetFileName = TargetName;

	}
	public String getSourceDir(){

	  return this.sourceDir;

	}
	public String getTargetFileName(){

		  return this.TargetFileName;

	}
	    public void initOutputFile(){

         try{

			 outputFile =  new File("excel_file_merger_log.txt");
			 if(outputFile.exists()) outputFile.delete();
			 outputFile.createNewFile();
             fileWriter =  new FileWriter(outputFile);
             printWriter =  new PrintWriter(fileWriter);

          }catch(Exception e){

		 					 e.printStackTrace();

		  }



	}

		 void processExcelFiles(){
			Iterator<File> fileItr  = this.fileSet.iterator();
			Map<String,String[][]> tempMap = new HashMap<String,String[][]>();
			 mySheet = mergedWorkbook.createSheet("MergedSheet");
			while (fileItr.hasNext()){
			    readExcelFile(fileItr.next().getAbsolutePath());
			}



			}
	 Map<String,String[][]> readExcelFile(String fileName) {

 		  String[][] excelData = null;
		  Map<String,String[][]>  excelSheetDataMap = new HashMap<String,String[][]>();
		  	String value = "";
			System.out.println("Processing: "+fileName);
			printWriter.println("Processing: "+fileName);
        if(fileName.endsWith("xlsx") && new File(fileName).length()<=(1024*1024)){
		try{
			  OPCPackage fis = OPCPackage.open(fileName);
			XSSFWorkbook   workbook = new XSSFWorkbook(fis);
			XSSFSheet[] sheets=  getXSSFSheets(workbook);
			int rowNum =0;
			int colNum = 0;
            ArrayList <XSSFSheet> usedSheets = new  ArrayList <XSSFSheet>();

			int sheetNum =sheets.length;
			int sheetCount =0;
			for(int k=0; k<sheetNum; k++){
						if(sheets[k].getRow(0) != null){
							++sheetCount;
							usedSheets.add(sheets[k]);
						}
			}
			for(int k=0; k<sheetCount; k++){

			rowNum = usedSheets.get(k).getLastRowNum()+1;
			excelFileData.clear();

			for (int i=0; i<rowNum; i++){

			if(sheets[k].getRow(i) != null){
				colNum = sheets[k].getRow(i).getLastCellNum();
			    excelData = new String[1][colNum];
				XSSFRow row = sheets[k].getRow(i);
					for (int j=0; j<colNum; j++){
						XSSFCell cell = row.getCell(j,XSSFRow.RETURN_BLANK_AS_NULL);
					    if(cell != null && cell.getCellType()!=XSSFCell.CELL_TYPE_BLANK) value =  DateUtil.isCellDateFormatted(cell)? cell.getDateCellValue():xssfCellToString(cell);
					    else value = null;
						excelData[0][j] = value;
						System.out.println("The value of cell ["+i+","+j+"] is: " + value);
						printWriter.println("The value of cell ["+i+","+j+"] is: " + value);
					}
						excelFileData.put(i,excelData );
					}
			   }
                            addDataToExcelFile(new File(fileName).getName()+"_sheet_"+k, excelFileData);
               }



		 fis.close();
			   }catch(Exception e){

				e.printStackTrace();
			   }

				}else if(fileName.endsWith("xls")&& new File(fileName).length()<=(1024*1024)){
			try{
                    				FileInputStream fis = new FileInputStream(fileName);
				  HSSFWorkbook   workbook = new HSSFWorkbook(fis);
				  HSSFSheet[] sheets= null;
	
                        sheets =getHSSFSheets(workbook);
						int rowNum =0;
						int colNum = 0;


			int sheetNum =sheets.length;
			int sheetCount =0;
			for(int k=0; k<sheetNum; k++){
						if(sheets[k].getRow(0) != null) ++sheetCount;
			}
			for(int k=0; k<sheetCount; k++){

                        rowNum = sheets[k].getLastRowNum()+1;
                        excelFileData.clear();
			for (int i=0; i<rowNum; i++){
			if(k>0 && rowNum!=0 ){
			if(sheets[k].getRow(i) != null) colNum = sheets[k].getRow(i).getLastCellNum();
			else colNum =0;
			excelData = new String[1][colNum];
				HSSFRow row = sheets[k].getRow(i);
				 if(row.getCell(3,HSSFRow.RETURN_BLANK_AS_NULL) !=null  && row.getCell(4,HSSFRow.RETURN_BLANK_AS_NULL)!=null && row.getCell(5,HSSFRow.RETURN_BLANK_AS_NULL) !=null   )
					for (int j=0; j<colNum; j++){
						HSSFCell cell = row.getCell(j,HSSFRow.RETURN_BLANK_AS_NULL);
					   if(cell != null && cell.getCellType()!=HSSFCell.CELL_TYPE_BLANK) value = HSSFDateUtil.isCellDateFormatted(cell)?df.format(cell.getDateCellValue()):hssfCellToString(cell);
					   else value = null;
						excelData[0][j] = value;
						System.out.println("The value is: " + value);
						printWriter.println("The value is: " + value);
					}
                                        excelFileData.put(i,excelData );
			}
			   }
                            addDataToExcelFile(new File(fileName).getName()+"_sheet_"+k, excelFileData);
                        }



		 fis.close();
			   }catch(Exception e){

				e.printStackTrace();
			   }

        }else{
            System.out.println("invalid file name, should be xls or xlsx or the file is too large >1MB");
            printWriter.println("invalid file name, should be xls or xlsx or the file is too large >1MB");
        }

	   return excelSheetDataMap;
  }
	String hssfCellToString (HSSFCell cell){

	int type = 0;
	Object result="";
	type = cell.getCellType();

		switch(type) {


		case 0://numeric value in excel
			result = cell.getNumericCellValue();
			break;
		case 1: //string value in excel
			result = cell.getStringCellValue();
			break;
		case 2: //boolean value in excel
			result = cell.getBooleanCellValue ();
			break;
		default:
			System.out.println("There are not support for this type of cell");
			}

	return result.toString();
	}
	String xssfCellToString (XSSFCell cell){

			int type =0;
			Object result =null;
			type = cell.getCellType();

    switch(type) {

    case 0://numeric value in excel
        result = cell.getNumericCellValue();
        break;
    case 1: //string value in excel
        result = cell.getStringCellValue();
        break;
    case 2: //boolean value in excel
        result = cell.getBooleanCellValue ();
        break;
    default:
    			System.out.println("There is no support for this type of cell");
        }

return result.toString();
}
	public void closeOutputFile(){

	         try{
				  outputFile = null;
				  printWriter.close();
				  fileWriter.close();

				 System.gc();

				 }catch(Exception e){

					 e.printStackTrace();

				}


	}

	void addDataToExcelFile(String sheetName, Map excelFileRowData) {
            int stringLen =sheetName.length();
            if (stringLen >=30){
            sheetName = sheetName.substring(0, 12)+"_"+sheetName.substring((stringLen-15), stringLen);

            }

			XSSFRow myRow = null;
			XSSFCell myCell = null;
				int mapSize = excelFileRowData.size();
				Set<Integer> excelKeys = excelFileRowData.keySet();
				Iterator excelItr = excelKeys.iterator();
				String[][] excelDataStrMap;
				int rows =0;
                        while(excelItr.hasNext()){

                         excelDataStrMap = (String[][]) excelFileRowData.get(excelItr.next());

			for (int rowNum = 0; rowNum < excelDataStrMap.length; rowNum++) {
				myRow = mySheet.createRow((globalRow));
				for (int cellNum = 0; cellNum < excelDataStrMap[rowNum].length; cellNum++) {
                    System.out.println("Adding Cell["+(globalRow)+","+cellNum+"]: "+excelDataStrMap[rowNum][cellNum]);
                    printWriter.println("Adding Cell["+(globalRow)+","+cellNum+"]: "+excelDataStrMap[rowNum][cellNum]);
						myCell = myRow.createCell(cellNum);
                        myCell.setCellValue( excelDataStrMap[rowNum][cellNum]);

				}
				++globalRow;
			}

                        excelDataStrMap =null;
                }

	}

	public void writeMergedFile(){

	String fileURL =this.getTargetFileName();
	if(!fileURL.endsWith(".xls") && !fileURL.endsWith(".xlsx")){
		fileURL+=".xlsx";
	}
	File processOutputFile = new File (fileURL);

	if(processOutputFile.exists()){
	    System.out.println("Deleting old output file: "+fileURL);
        printWriter.println("Deleting old output file: "+fileURL);
		processOutputFile.delete();
	}
		    System.out.println("Saving new output file: "+fileURL);
        printWriter.println("Saving new output file: "+fileURL);
		try {
			FileOutputStream out = new FileOutputStream(fileURL);
			mergedWorkbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
		public static void showUsage(){

			System.out.println("Usage:  ExcelFileMergerRecon -s 'C:\\temp' -t 'C:\\users\\public'");
			System.out.println("Where: ");
			System.out.println("s: Source folder of Excel files");
			System.out.println("t: Target file for Merged Excel files");

     }
	  HSSFSheet[] getHSSFSheets(HSSFWorkbook workbook){
			int size =workbook.getNumberOfSheets();
	      HSSFSheet[] sheets = new HSSFSheet[size];

            for(int i=0; i<size;i++){
			sheets[i]=workbook.getSheetAt(i);
			}
			return sheets;
	 }
	  XSSFSheet[] getXSSFSheets(XSSFWorkbook workbook){
	      int size =workbook.getNumberOfSheets();
	      XSSFSheet[] sheets = new XSSFSheet[size] ;

            for(int i=0; i<size;i++){
			sheets[i]=workbook.getSheetAt(i);
			}
			return sheets;
	 }
    public void getExcelFiles(){

		  try{

			File excelSourceFolder = new File(this.sourceDir);
			printWriter.println("Reading files in: "+excelSourceFolder.getAbsolutePath());
			if(excelSourceFolder.exists()){

				 excelFiles = excelSourceFolder.listFiles();
				 int fileCount = excelFiles.length;

				 printWriter.println("\n\nGetting file list...");
				 System.out.println("\n\nGetting file list...");

				 for(int i=0; i<fileCount; i++){

					this.fileSet.add(excelFiles[i]);

				 }
				   printWriter.println(this.fileSet);
				   printWriter.println(this.fileSet.size()+" files found.");
				   System.out.println("\n\n"+this.fileSet);
				   System.out.println(this.fileSet.size()+" files found.");

		  }else{

				printWriter.println(excelSourceFolder.getName()+" does not exist!");
				printWriter.println("exiting...");
				System.out.println(excelSourceFolder.getName()+" does not exist!");
				System.out.println("exiting...");
				closeOutputFile();
				System.exit(0);

		  }
	   }catch(Exception e){

			        e.printStackTrace();
					printWriter.println("Exiting...");
				    System.out.println("Exiting...");
					closeOutputFile();
					System.exit(0);


		  }

	}

	public static void main(String [] args){

				            String source = "";
				            String target = "";

		if(args.length==0){
					 System.out.print("Using default paramaters...");
				     source = "resources";
					 target = "results\\merged_excel_file.xlsx";

	     }else{

				int argsCount = args.length;
				try{
					for(int i=0; i< argsCount; i++){
									args[i] = args[i].trim();
								if(args[i].equalsIgnoreCase("-s")){
									source =args[i+1].trim();
									source =source.replace("\\","\\\\");
									source =source.replace("\'","");
									source =source.replace("\"","");
								   System.out.println("\nSource: "+source );

								 }else 	if(args[i].equalsIgnoreCase("-t")){
									target =args[i+1].trim();
									target =target.replace("\\","\\\\");
									target =target.replace("\'","");
									target =target.replace("\"","");
                                     System.out.println("\nTarget:"+ target );

								 }


					   }
								 if(source.isEmpty() && target.isEmpty()) {
									showUsage();
								 }else if(source.isEmpty() ){
									 System.out.print("\nInvalid source folder specified");
									 System.out.print("\nUsing default source folder...");

								 }else if(target.isEmpty() ){
									 System.out.print("\nInvalid target folder specified");
									 System.out.print("\nUsing default target folder...");

								 }
										}catch(Exception e){
												   e.printStackTrace();
												   showUsage();

											}

			}

				   System.out.println("\nSource Folder: "+source);
				   System.out.println("\nTarget File: "+target);

				   new ExcelFileMergerRecon( source,  target);
				   System.out.println("\n\nProcess Complete.");
				   System.out.println("\nExiting...");

		  }

	}