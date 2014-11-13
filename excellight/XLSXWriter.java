package excellight;

import java.io.File;


import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.nio.file.FileAlreadyExistsException;
import java.text.NumberFormat;
import java.util.ArrayList;

/**
 * Simple and light-weight java class for writing raw data to .xlsx excel files.
 * <p>
 * Styling options are unavailable however they can be added manually in the final product.
 * <p>
 * Data is subject to multiple files which creates more opportunities for I/O errors.
 * For this reason, exception handling is left to the user.
 * 
 * @author Thomas Burghardt
 * @version 0.0
 */
public class XLSXWriter {

	public final static String TEMPLATE_FILE_NAME="excel_template";
	
	private boolean saved;
	private final String FILE_NAME;
	private ArrayList<String> sharedStrs;
	private int sharedStrsCount=0;
	private ArrayList<String> sheets;
	private PrintWriter sharedStrings;
	private PrintWriter contentTypes;
	private PrintWriter app;
	private PrintWriter workBookRels;
	private PrintWriter workBook;
	private NumberFormat doubleFormatter;
	private NumberFormat intFormatter;
	private static final String FILE_SEP="/";
	
	
	/**
	 * The constructor opens the file according to the name given, however the extension is implied so can be left out.
	 * The file name given is final. There should be one XLSXWriter object per file written and the class is intended as a simple one time writer.
	 * <p>
	 * Once the object has been saved, it is no longer of any use.
	 * 
	 * @param fName filename or full path of file to be written
	 * @throws IOException related to if a directory with the argument provided already exists
	 */
	public XLSXWriter(String fName) throws IOException{
		
		//check and remove extension if it exists
		if(fName.substring(fName.length()-5,fName.length()).equals(".xlsx"))
			FILE_NAME = fName.substring(0,fName.length()-5);
		else
			FILE_NAME = fName;
		
		File f = new File(FILE_NAME+".xlsx");//check if file exists
		if(f.exists()) {throw new FileAlreadyExistsException("There already exists a file with the given name."); }
		
		saved=false;
		createDocument();
		sharedStrs=new ArrayList<String>();
		sheets=new ArrayList<String>();
		
		//print writers for creating the necessary XML docs
		contentTypes = new PrintWriter(new FileOutputStream(new File(FILE_NAME+"//[Content_Types].xml"),true));
		app = new PrintWriter(new FileOutputStream(new File(FILE_NAME+"//docProps//app.xml"),true));
		workBookRels = new PrintWriter(new FileOutputStream(new File(FILE_NAME+"//xl//_rels//workbook.xml.rels"),true));
		workBook = new PrintWriter(new FileOutputStream(new File(FILE_NAME+"//xl//workbook.xml"),true));
		sharedStrings = new PrintWriter(new FileOutputStream(new File(FILE_NAME+"//xl//sharedStrings.xml"),true));
		
	}
	/*
	* creates the directory containing all the XML files and sub-directories that make up an excel file
	* this is done from a template directory that contains the base information of each file
	*/
	private void createDocument() throws IOException{
		InputStream resourceStream = XLSXWriter.class.getResourceAsStream(TEMPLATE_FILE_NAME+".zip");
		ZipFileManager.unzip(resourceStream, FILE_NAME);
		
	}
	/**
	 * Add a worksheet to the excel document.
	 * <p>
	 * If the document has been saved it can no longer be accessed. 
	 * This will throw an IllegalStateException (which is a RunTimeEception).
	 * 
	 * @param data two dimensional string-array containing the data to be written, values that can be interpreted as integers or doubles will be
	 * @param sheetName the name of the worksheet to be written
	 * @throws IOException due to errors in initialization or corrupted files
	 */
	public void writeSheet(String[][] data, String sheetName) throws IOException {
		
		if(saved)
			throw new IllegalStateException("This document has already been saved. Additional sheets can no longer be added.");

		//sheet name checking
		if(sheetName==null)
			sheetName="Sheet"+sheets.size();
		else{
			int sheetNum=1;
			String sheetTemp=sheetName;
			while(sheets.contains(sheetTemp)){
				sheetTemp=sheetName+"("+sheetNum+")";
				sheetNum++;
			}
			sheetName=sheetTemp;
		}
		//find largest column length
		int maxCol=0;
		for(int i=0;i<data.length;i++){
			if(data[i].length>maxCol)
				maxCol=data[i].length;
		}
		//create sheet
		Sheet sheet = new Sheet(this,sheetName,data.length,maxCol);
		for(int i=0;i<data.length;i++)
			sheet.addRow(data[i]);
		sheet.finalizeSheet();
	}
	
	/*
	 * Finalize the document.
	 * <p>
	 * All closing XML tags and necessary information will be added to the XML files that makes up the excel document.
	 */
	private void finalizeDoc(){
		
		
		app.print(sheets.size()+"</vt:i4></vt:variant></vt:vector></HeadingPairs>" +
				"<TitlesOfParts><vt:vector size=\"" +sheets.size()
		+"\" baseType=\"lpstr\">");
		
		int sheetNum=1;
		for(String sheetName : sheets){
			app.println( "<vt:lpstr>"+sheetName+"</vt:lpstr>");
			
			workBook.print("<sheet name=\""+sheetName+"\" sheetId=\""+sheetNum+"\" r:id=\"rId"+sheetNum+"\"/>");
			
			
			contentTypes.println("<Override PartName=\"/xl/worksheets/sheet"+sheetNum+".xml\" "+
			        "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
			
			workBookRels.println("<Relationship Target=\"worksheets/sheet"+sheetNum+".xml\" " +
					"Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" " +
					"Id=\"rId"+sheetNum+"\"/>");
			
			sheetNum++;
		}
		contentTypes.print("</Types>");
		
		app.print("</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate>" +
				"<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>" +
				"<AppVersion>14.0300</AppVersion></Properties>");
		
		workBook.print("</sheets><calcPr calcId=\"145621\"/></workbook>");
		
		String[] type = { "theme","styles","sharedStrings"};
		String[] target = {type[0]+"//"+type[0]+"1",type[1],type[2]}; // uh oh//?
		for(int i=0;i<type.length;i++)
			workBookRels.print("<Relationship  Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"+type[i]+"\" " +
					"Target=\""+target[i]+".xml\" Id=\"rId"+(sheets.size()+i+1)+"\"/>");
		workBookRels.print("</Relationships>");
		
		sharedStrings.println("count=\""+sharedStrsCount+"\" uniqueCount=\""+sharedStrs.size()+"\">");
		for(String shared : sharedStrs)
			sharedStrings.println("<si><t>"+shared+"</t></si>");
		sharedStrings.print("</sst>");
		
	}

	/**
	 * Saves the file, closes all the writers and zips the contents of the newly created directory
	 * <p>
	 * Saving twice is redundant so attempting to save again throws an IllegalStateException (which is a RunTimeEception).
	 */
	public void save(){

		if(saved)
			throw new IllegalStateException("This document has already been saved. Cannot save again.");
		//close writers
		finalizeDoc();
		sharedStrings.close();
		contentTypes.close();
		app.close();
		workBookRels.close();
		workBook.close();
		
		//zip file and add XLSX file extension
		ZipFileManager zipper = new ZipFileManager(FILE_NAME,FILE_NAME+".zip");
		zipper.zip();
		File excelDoc = new File(FILE_NAME+".zip");
		File name = new File(FILE_NAME+".xlsx");
		excelDoc.renameTo(name);
		
		saved=true;
	}
	/**
	 * Number formats for doubles can be specified to reflect how many decimal places they should have.
	 * 
	 * @param nFormat the number format object thats specifies how doubles should appear in the file
	 */
	public void specifyDoubleFormat(NumberFormat nFormat){
		doubleFormatter  = nFormat;
	}
	/**
	 * Number formats for integers can be specified to reflect how many decimal places they should have.
	 * 
	 * @param nFormat the number format object thats specifies how integers should appear in the file
	 */
	public void specifyIntFormat(NumberFormat nFormat){
		intFormatter  = nFormat;
	}
	/**
	 * Intended to be a helper method however, it may come in handy.
	 * 
	 * @return the number of sheets currently in the file
	 */
	public int numSheets(){
		return sheets.size();
	}
	/**
	 * Returns the file name (not the template file name but the one being created).
	 * 
	 * @return the name of the file
	 */
	public String getFileName(){
		return FILE_NAME;
	}
	/**
	 * Saving twice is redundant and will result in an error message. This method can be used to check if the file has already been saved.
	 * 
	 * @return a boolean that specifies weather the file has been saved yet 
	 */
	public boolean isSaved(){
		return saved;
	}

	
	//Private helper class that holds worksheet information related to XML tags and file contents.
	private class Sheet {
		
		private XLSXWriter mDoc;
		private PrintWriter sheetWriter;
		private int numRows=0;
		public static final String EOF="</sheetData>" +
				"<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" +
				"</worksheet>";
		public static final String POST_DIM = "\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>" +
				"<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/><sheetData>";
		
		public Sheet(XLSXWriter doc,  String sheetName, int row, int cols) throws IOException{
			mDoc=doc;
			mDoc.sheets.add(sheetName);
			int sheetNum=mDoc.numSheets();
			File mFile = new File(doc.getFileName()+FILE_SEP+"xl"+FILE_SEP+"worksheets"+FILE_SEP+"sheet"+sheetNum+".xml");
			InputStream resourceStream = XLSXWriter.class.getResourceAsStream(TEMPLATE_FILE_NAME+".zip");
			if(sheetNum>1)
				ZipFileManager.copyZipFile(resourceStream, mFile);
			sheetWriter = new PrintWriter(new FileOutputStream(mFile,true));
			sheetWriter.println(ExcelUtility.excelIntToCol(cols)+row+POST_DIM);
		}
		public void finalizeSheet(){
			sheetWriter.print(EOF);
			sheetWriter.close();
		}
		
		
		private String rowStart(int rowNum, int numCols){
			return "<row r=\""+rowNum+"\" spans=\"1:"+numCols+"\" x14ac:dyDescent=\"0.25\">";
		}
		private String rowEnd(){
			return "</row>";
		}
		//add a single row from an array of data to the current worksheet
		public void addRow(String[] data){
			
			numRows++;
			String row=rowStart(numRows,data.length);
			for(int i=0;i<data.length;i++){
				if(data[i]==null || data[i].equals(""))
					continue;
				row+="<c r=\""+ExcelUtility.excelIntToCol(i+1)+numRows+"\"";
				//handle different data types
				if(ExcelUtility.isInt(data[i])){
					if(intFormatter!=null)
						row+="><v>"+intFormatter.format(Integer.parseInt(data[i]));
					else
						row+="><v>"+Integer.parseInt(data[i]);
				}
				else if(ExcelUtility.isDouble(data[i])){
					if(doubleFormatter!=null)
						row+="><v>"+doubleFormatter.format(Double.parseDouble(data[i]));
					else
						row+="><v>"+Double.parseDouble(data[i]);
						
				}
				//is a shared string
				else{
					mDoc.sharedStrsCount++;
					row+=" t=\"s\"><v>";
					ArrayList<String> shared = mDoc.sharedStrs;
					if(shared.contains(data[i]))
						row+=shared.indexOf(data[i]);
					else{
						row+=shared.size();
						shared.add(data[i]);
					}
				}
				row+="</v></c>";
			}
			row+=rowEnd();
			sheetWriter.println(row);		
		}
		
	}
}
