package excellight;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Simple and light-weight java class for the raw data reading .xlsx documents.
 * <p>
 * Other information, such as styling, can be present in the document but it will be ingored.
 * <p>
 * Data is subject to multiple files which creates more opportunities for I/O errors.
 * For this reason, exception handling is left to the user.
 * 
 * @author Thomas Burghardt
 * @version 0.0
 */
public class XLSXReader {
	
	private final String FILE_NAME;
	private HashMap<String,Integer> worksheets;
	private static final String FILE_SEP="/";
	
	/**
	 * The constructor opens the file according to the name given, however the extension is implied so can be left out.
	 * 
	 * @param fName file name or full path of file to be read
	 * @throws IOException these exceptions are related to if the file does not exist, or the file path or type is incorrect
	 */
	public XLSXReader(String fName) throws IOException{
		//check and remove extension if it exists
		if(fName.substring(fName.length()-5,fName.length()).equals(".xlsx"))
			FILE_NAME = fName;
		else
			FILE_NAME = fName+".xlsx";
		
		initializeWorksheets();
	}
	/*
	 * Finds the shared strings of the excel document and constructs a map for reading
	 * the file containing shared strings is ordered, the integer corresponding to the the string is the string's position in the file
	 */
	private HashMap<Integer,String> sharedStrings(String fileContents){
		
		HashMap<Integer,String> stringMap = new HashMap<Integer,String>();
		Pattern p = Pattern.compile("<si>.*?<t.*?>(.*?)</t>.*?</si>"); //regular expression that finds values of shared strings

		String[] sharedStrs = ExcelUtility.findMatches(p,fileContents);
		for(int i=0;i<sharedStrs.length;i++)
			stringMap.put(i, sharedStrs[i]);
		return stringMap;
	}
	/*
	 * private method to get the names of the worksheets and map them to their number
	 * helper method to initializeWorksheets
	 */
	private HashMap<String,Integer> worksheetMap(String fileContents){
		
		HashMap<String, Integer> stringMap = new HashMap<String, Integer>();
		Pattern p = Pattern.compile("<sheet.*?name=\"(.*?)\".*?/>");//regular expression to find the name of a worksheet

		String[] worksheetMacthes = ExcelUtility.findMatches(p,fileContents);
		for(int i=0;i<worksheetMacthes.length;i++)
			stringMap.put(worksheetMacthes[i], i+1);
		return stringMap;
	}
	
	//open the workbook XML file and initialize the map holding the worksheets and their corresponding numbers
	private void initializeWorksheets() throws IOException{
		File file = new File(FILE_NAME);
		ZipFile zipFile = new ZipFile(file);
		ZipEntry workbookZip =  zipFile.getEntry("xl"+FILE_SEP+"workbook.xml");
		
		InputStream input = zipFile.getInputStream(workbookZip);
		
		//read workbook and map worksheet names to their respective numbers
		BufferedReader br = new BufferedReader(new InputStreamReader(input, "UTF-8"));
		String workbook = ExcelUtility.getFileContents(br);
		worksheets = worksheetMap(workbook);
		zipFile.close();
	}
	/*
	 * interprets the fContents as a worksheet file found in xl/worksheets and returns the data contained within
	 */
	private static String[][] readWorksheetXML(String fContents, HashMap<Integer,String> sharedStrs){
		
		/*
		 * forming the regular expression to handle comma-delimited data -> separating values is non-trivial
		 * split into two main groups: sharedStrings and notShared
		 * each subdivided into three groups:
		 * 
		 * 1. column Id
		 * 2. row
		 * 3. value
		 * 
		 * (sharedStrings gas 4 sub-groups but the third is the expression signifying it is a sharedString, so is skipped
		 */
		String regVal = "<v>(.*?)</v>";
		String regRow = "<c r=\"([A-Z]*?)(\\d*?)\"";
		String regCell = regRow+"[^/]*?>.*?"+regVal;
		String regCellShared = regRow+"\\s*?(?:s=\"\\d*?\")??\\s*?(t=\"s(?:tr)??\").*?"+regVal;
		String finalReg="(?:"+"(?:"+regCellShared+")"+"|"+"(?:"+regCell+")"+")";
		Pattern valPattern = Pattern.compile(finalReg);
		Pattern dimPattern = Pattern.compile("<dimension ref=\"[A-Z]*?\\d*?:([A-Z]*?)(\\d*?)\"");
		
		//find number of columns
		Matcher dimMatcher = dimPattern.matcher(fContents);
		int numRows=0;
		int numCols=0;
		if(dimMatcher.find());{
			numRows=Integer.parseInt(dimMatcher.group(2));
			numCols=ExcelUtility.excelColToInt(dimMatcher.group(1));
		}
		String[][] data = new String[numRows][numCols];
		Matcher valMatcher = valPattern.matcher(fContents);
		int ind=0;
		
		//find the matches and add values
		while (valMatcher.find(ind)) {
			
			if(valMatcher.group(1)!=null){ //value is a shared string
				int row=Integer.parseInt(valMatcher.group(2));
				int col=ExcelUtility.excelColToInt(valMatcher.group(1));
				String val = valMatcher.group(4);
				data[row-1][col-1]=
						valMatcher.group(3).equals("t=\"s\"") ? 
								sharedStrs.get(Integer.parseInt(val)) : val;
				ind=valMatcher.end(4);
			}
			else{//value is not a shared string (likely a number)
				
				int row=Integer.parseInt(valMatcher.group(6));
				int col=ExcelUtility.excelColToInt(valMatcher.group(5));
				data[row-1][col-1]=valMatcher.group(7);
				ind=valMatcher.end(7);
			}
	    }
		return data;
	
	}
	/**
	 * Reads a worksheet specified by its index within the excel document.
	 * 
	 * @param sheetNum the number of the sheet to collect data from
	 * @return two dimensional array containing the excel data, empty values are null
	 * @throws IOException should not occur, but is possible if initial data is incorrect
	 */
	public String[][] readSheet(int sheetNum) throws IOException{
		File file = new File(FILE_NAME);
		String strFile="";
		String sheetFile="";
		ZipFile zipFile = new ZipFile(file);
		ZipEntry sheet =  zipFile.getEntry("xl"+FILE_SEP+"worksheets"+FILE_SEP+"sheet"+sheetNum+".xml");
		ZipEntry sharedStrs =  zipFile.getEntry("xl"+FILE_SEP+"sharedStrings.xml");
		
		//reed the sheet
		InputStream input = zipFile.getInputStream(sheet);
		BufferedReader br = new BufferedReader(new InputStreamReader(input, "UTF-8"));
		sheetFile = ExcelUtility.getFileContents(br);
		
		//reed the shared strings
		InputStream input2 = zipFile.getInputStream(sharedStrs);
		BufferedReader br2 = new BufferedReader(new InputStreamReader(input2, "UTF-8"));
		strFile = ExcelUtility.getFileContents(br2);
			
		zipFile.close();
		
		HashMap<Integer,String> sharedStrings = sharedStrings(strFile);
		return readWorksheetXML(sheetFile,sharedStrings);
	}
	/**
	 * Reads a worksheet specified by its name within the excel document. 
	 * 
	 * @param sheetName the name of the sheet to collect data from
	 * @return two dimensional array containing the excel data, empty values are null
	 * @throws IOException should not occur, but is possible if initial data is incorrect
	 */
	public String[][] readSheet(String sheetName) throws IOException{
		int sheetNum = worksheets.get(sheetName);
		return readSheet(sheetNum);
	}
	/**
	 * Returning the number of worksheets allows indexing in the readSheet method.
	 * 
	 * @return the number of worksheets in this file
	 */
	public int numSheets(){
		return worksheets.size();
	}
	/**
	 * Returns a string-array that can be used to pass to the readSheet method.
	 * 
	 * @return an string-array containing the names of the worksheets in the file
	 */
	public String[] sheetNames(){
		return worksheets.keySet().toArray(new String[numSheets()]);
	}
}
