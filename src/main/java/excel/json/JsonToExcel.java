package excel.json;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.util.HashMap;
import java.util.Map.Entry;
import java.util.logging.Logger;

import com.metamug.parser.schema.Request;
import com.metamug.parser.schema.Resource;
import com.metamug.parser.service.OpenAPIParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class JsonToExcel {

	public static void exportswaggerUrlToExcel(String swaggerUrl) throws Exception{
		// TODO Auto-generated method stub

		XSSFWorkbook workbook = new XSSFWorkbook();
        
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("API Details");

        String jsonString = readUrl(swaggerUrl);
//        Logger.getLogger(JsonToExcel.class.getName()).info(jsonString);


        int count = 0;
        int rownum = 0;

        Row row = sheet.createRow(rownum++);
        String [] objArr = new String[]{"S.No.","API_URL","CONTROLLER", "DESCRIPTION", "RESPONSE_TYPE"};
        int cellnum = 0;
        
        for (String obj : objArr){
           Cell cell = row.createCell(cellnum++);
                cell.setCellValue((String)obj);
            
        }
        
        rownum++;

        OpenAPIParser parser = new OpenAPIParser(jsonString);

        for (Entry<String, Resource> e : parser.getBackend().getResourceList().entrySet()) {
        	
        	String key = e.getKey();
            Resource resource = e.getValue();
            System.out.println();

            for(Request request : resource.getRequest()){
                row = sheet.createRow(rownum++);
                Cell cell = row.createCell(0);
                cell.setCellValue(key);

                cell = row.createCell(1);
                cell.setCellValue(String.valueOf(request.getMethod()));

                cell = row.createCell(2);
                cell.setCellValue(String.valueOf(request.getDescString()));

            }

        
        }
        
        System.out.println(count);
       
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("API_DETAILS.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("API_DETAILS.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
	
	private static String readUrl(String urlString) throws Exception {
	    BufferedReader reader = null;
	    try {
	        URL url = new URL(urlString);
	        reader = new BufferedReader(new InputStreamReader(url.openStream()));
	        StringBuffer buffer = new StringBuffer();
	        int read;
	        char[] chars = new char[1024];
	        while ((read = reader.read(chars)) != -1)
	            buffer.append(chars, 0, read); 

	        return buffer.toString();
	    } finally {
	        if (reader != null)
	            reader.close();
	    }
	}
	

}
