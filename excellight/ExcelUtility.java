package excellight;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Utility methods for the excelIO classes
 * @hide
 */
public class ExcelUtility {

	//integer test
	public static boolean isInt(String s) {
		if (s == null)
			return false;
		try {
			Integer.parseInt(s);
		} catch (NumberFormatException e) {
			return false;
		}
		return true;
	}
	//double test
	public static boolean isDouble(String s) {
		if (s == null)
			return false;
		try {
			Double.parseDouble(s);
		} catch (NumberFormatException e) {
			return false;
		}
		return true;
	}
	/*
	 * takes an integer and maps it to an excel column id
	 * i.e. 1 becomes A1, 2 becomes A2,...
	 */
	public static String excelIntToCol(int num) {

		String result = "";
		while (num > 0) {
			num--;
			int remainder = num % 26;
			char digit = (char) (remainder + 65);
			result = digit + result;
			num = (num - remainder) / 26;
		}
		return result;
	}
	/*
	 * takes an excel column ID and maps it to an integer, only accepts capital letters
	 * i.e. A1 becomes 1, A2 becomes 2,...
	 */
	public static int excelColToInt(String col){
		char[] colLetters = col.toCharArray();
		int result=0;
		for(int i=0;i<colLetters.length;i++){
			int letter = (int)(colLetters[i]-64);
			if(letter<1||letter>26)
				return -1; //not a capital letter
			result+=letter*Math.pow(26, colLetters.length-(i+1));
		}
		return result;
	}
	/*
	 * copy file:
	 * Important to copy the template files so they can be altered in the copy but the original can still be read from
	 */
	public static void copyFile( File from, File to ) throws IOException {
	    Files.copy( from.toPath(), to.toPath() );
	}
	//for creating blank excel directory to be filled and zipped
	public boolean createBlankDocument(String docName){
		return (new File("new_excel")).mkdirs();
	}
	/*
	 * returns entire contents of file in a string
	 * important for regular expression evalutaion
	 */
	public static String getFileContents(BufferedReader reader){
		String fContents="";
		String line;
		try {
			while((line=reader.readLine())!=null){
				fContents += line;
			}
			reader.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return fContents;
	}
	/*
	 * search regular expression and return array of matches
	 * can be used with getFile contents to isolate macthes in entire file
	 */
	public static String[] findMatches(Pattern p, String s){
		Matcher matcher = p.matcher(s);
		int ind=0;
		ArrayList<String> matches = new ArrayList<String>();
		
		while (matcher.find(ind)) {
	
			for(int i=1;i<matcher.groupCount()+1;i++){
				if(matcher.group(i)!=null){
					ind=matcher.end(i); //makes the matcher backs up just after the previous match to continue search
					matches.add(matcher.group(i));
				}
			}
	    }
		return matches.toArray(new String[0]);
	}
}
