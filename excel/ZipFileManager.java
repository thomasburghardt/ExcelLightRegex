package excellight;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileAlreadyExistsException;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * Convenience class for zipping files
 * SRC folder is deleted after completion
 * @hide
 */
public class ZipFileManager {

	private List<String> fileList;
	private List<String> dirList;
    private final String OUTPUT_ZIP_FILE;
    private final String SOURCE_FOLDER;
    private File SOURCE_FILE;
    private static final String FILE_SEP="/";

    //given a source folder, outputs the zip file to the target specified
    public ZipFileManager(String source, String target) {
    	OUTPUT_ZIP_FILE=target;
    	SOURCE_FOLDER=source;
    	fileList = new ArrayList<String>();
    	dirList = new ArrayList<String>();
    	SOURCE_FILE=new File(SOURCE_FOLDER);
    }
    //finalize zip document
    public void zip() 
    {
    	generateFileList(SOURCE_FILE);
    	zipIt(OUTPUT_ZIP_FILE);
    	//deletes old src folder
    	for (String dir : this.dirList) {
    		File deleteMe = new File(dir);
    		deleteMe.delete();
    	}
    }

    /*
     * actual zip process
     * Uses outputstreams and input streams to read and write files
     */
    private void zipIt(String zipFile){
 
     byte[] buffer = new byte[1024];
     
     try{
    	File zip = new File(zipFile);
    	FileOutputStream fos = new FileOutputStream(zip);
    	ZipOutputStream zos = new ZipOutputStream(fos);
 
    	//loop through file list create by generateFileList method
    	for(String file : this.fileList){

    		ZipEntry ze= new ZipEntry(file);
        	zos.putNextEntry(ze);
        	
        	FileInputStream in =  new FileInputStream(SOURCE_FOLDER + FILE_SEP + file);
        	
        	//write bytes of file to output stream
        	int len;
        	while ((len = in.read(buffer)) > 0) {
        		zos.write(buffer, 0, len);
        	}
        	in.close();
        	//add contents that need to be deleted when zipping is complete
        	File deleteMe = new File(SOURCE_FOLDER + File.separator + file);
        	if(!deleteMe.isDirectory())
        		deleteMe.delete();

        	zos.closeEntry();
    	}
    	
    	zos.close();
    }catch(IOException ex){
       ex.printStackTrace();   
    }
   }
    //copy a single zip file to a desired location
   public static void copyZipFile(InputStream resourceStream,File outputFile){
	   byte[] buffer = new byte[1024];
	   
	   ZipInputStream zis = new ZipInputStream(resourceStream);
	   try {
		   ZipEntry entry= zis.getNextEntry();
		   if(entry==null || entry.isDirectory()){
			   System.err.println("Error copying zip file."); //file not found or it is a directory
			   //in the case of a directory use unzip
			   return;
		   }
		   FileOutputStream fos = new FileOutputStream(outputFile);             

		   int len;
		   while ((len = zis.read(buffer)) > 0) {
			   fos.write(buffer, 0, len);
		   }
	
		   fos.close();  
	   } catch (IOException e) {
		e.printStackTrace();
	   }
	   
   }
   //unzip a directory to a target out folder
   public static void unzip(InputStream resourceStream,String outputFolder) throws FileAlreadyExistsException{
	   byte[] buffer = new byte[1024];
	   
	   //create output directory if not exists
	   File folder = new File(outputFolder);
	   if(!folder.exists())
		   folder.mkdirs();
	   else
		   throw new FileAlreadyExistsException("A folder with the specified file name already exists.");

	   ZipInputStream zis = new ZipInputStream(resourceStream);
	   ZipEntry entry;

	   try {
		//loop through directory contents
		while((entry = zis.getNextEntry()) != null) {

			   if(entry.isDirectory()){
				   new File(outputFolder + FILE_SEP + entry.getName()).mkdirs();
				   continue;
			   }
			   else{
				   File newFile = new File(outputFolder + FILE_SEP + entry.getName());
				   FileOutputStream fos = new FileOutputStream(newFile);             
	
				   int len;
				   //read file
				   while ((len = zis.read(buffer)) > 0) {
					   fos.write(buffer, 0, len);
				   }
				   fos.close();   
			   }
		 }
		 zis.closeEntry();
		 zis.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	   

   }
   //recursive method to generate a file list given a node
   private void generateFileList(File node){
		if(node.isFile())
			fileList.add(generateZipEntry(node.getPath().toString()).replaceAll("\\\\", FILE_SEP)); //replace ensures correct file separation
		
		//node is a directory so  all sub contents need to be added to the list
		if(node.isDirectory()){
			
			String[] subNote = node.list();
			
			for(String filename : subNote)
				generateFileList(new File(node, filename));
			dirList.add(node.getPath().toString());
		}
 
    }
   
    private String generateZipEntry(String file){
    	return file.substring(SOURCE_FOLDER.length()+1, file.length());
    }
}
