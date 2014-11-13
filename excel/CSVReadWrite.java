package excellight;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.LineNumberReader;
import java.io.PrintWriter;
import java.nio.file.FileAlreadyExistsException;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Simple light-weight java class for writing raw data to .csv excel files. It behaves as a static class and cannot be sub-classed.
 * <p>
 * This class is only for <code>comma-delimited</code> excel files. Note that .csv (comma-separated values) is already a 
 * vaguely specified file format. Some files with .csv extensions may not actually be comma-separated. It is up to the user to determine
 * where it is appropriate to use this class, and it is recommended that the files provided come from a trusted source.
 * <p>
 * Accuracy of final data is not guaranteed, but all tests have so far proven it's robustness.
 * <p>
 * Since data is contained in a single file, all exceptions are handled.
 * 
 * @author Thomas Burghardt
 * @version 0.0
 */
public final class CSVReadWrite {
	/*
	 * the large regular expression to weed out separate values, appears to be robust
	 * works on values such as:
	 * 
	 * ","
	 * "",","
	 * ext.
	 */
	private final static String postComma="(?:[^\"]|\"[^,])";
	private final static String comma=","+postComma;
	private final static String after="(?:"+comma+"|$)";
	private final static String regA = "(?:(?:,|^)\"(.*?)\""+after+")";
	private final static String regB = "(?:(?:,|^)([^\",]*?.*?[^\"]*?)(?:,|$))";
	private final static String regexCSV ="(?:"+regA+"|"+regB+")";
	
	//prevent subclassing
	private CSVReadWrite(){
		throw new AssertionError("CSVReadWrite cannot be subclassed.");
	}
	//get number of lines in the excel file
	private static int getNumLines(String fName){
		int result=-1;
		try {
			LineNumberReader  lnr = new LineNumberReader(new FileReader(fName));
			lnr.skip(Long.MAX_VALUE);
			result = lnr.getLineNumber();
			lnr.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return result;
	}
	/**
	 * Read a single .csv excel file.
	 * 
	 * @param fName the name of the file  or full path
	 * @return  two dimensional array containing the data. Empty values will be null
	 * @throws FileNotFoundException 
	 */
	public static String[][] read(String fName) throws FileNotFoundException{
		
		//check and add extension if it doesn't exists
		if(!fName.substring(fName.length()-4,fName.length()).equals(".csv"))
			fName += ".csv";
				
		Pattern p = Pattern.compile(regexCSV);
		String[][] data = new String[getNumLines(fName)][];
		
		Scanner reader = new Scanner(new FileReader(fName));
		int rowNum=0;
		while(reader.hasNextLine()){
			String line = reader.nextLine();
				//handle corrupting characters
				line=line.replaceAll("\"\"", "\"");
				
				Matcher matcher = p.matcher(line);
		    	int ind=0;
		    	ArrayList<String> row = new ArrayList<String>();
		    	//loop through matches and read data accordingly
		    	while (matcher.find(ind)) {

		    		for(int i=1;i<matcher.groupCount()+1;i++){
		    			if(matcher.group(i)!=null){
		    				ind=matcher.end(i);
		    				row.add(matcher.group(i));
		    			}
		    		}
		        }
		    	data[rowNum]=row.toArray(new String[0]);
		    	rowNum++;
			}
			reader.close();
		return data;
	}
	
	/**
	 * Write a single .csv excel file.
	 * 
	 * @param data two dimensional array containing the data. null values will be empty
	 * @param fName the name of the file or full path
	 * @throws FileAlreadyExistsException
	 * @throws FileNotFoundException 
	 */
	public static void write(String[][] data, String fName) throws FileAlreadyExistsException, FileNotFoundException{
		//check and remove extension if it exists
		if(fName.substring(fName.length()-4,fName.length()).equals(".csv"))
			fName = fName.substring(0,fName.length()-4);
		
		File f = new File(fName+".csv");//check if file exists
		if(f.exists()) {throw new FileAlreadyExistsException("There already exists a file with the given name."); }
		
		PrintWriter printer = new PrintWriter(fName+".csv");
		for(int i=0;i<data.length;i++){
			for(int j=0;j<data[i].length;j++){
					
				String cellValue=data[i][j];
				if(cellValue==null)
					cellValue="";
				//check whether the value contains corrupting characters
				if(cellValue.contains("\"")||cellValue.contains(",")){
					cellValue=cellValue.replaceAll("\"","\"\"");
					cellValue="\""+cellValue+"\"";
				}
				printer.print(cellValue+",");
			}
			printer.println();
		}
		printer.close();	
	}
}
