package rpa;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import com.aspose.cells.Cell;
import com.aspose.cells.CellValueType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import okhttp3.MediaType;
import okhttp3.MultipartBody;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;
import utilities.ExcelReader;


public class ExcelDownloader {
	String token="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiRW1wbG95ZWUsQWNjb3VudCBNYW5hZ2VyLFN1cGVydmlzb3IsUmVhZCBPbmx5LFJ1bGUgRW5naW5lIiwiZnVsbG5hbWUiOiJEYW55YWwgQW1hbiwgTXVoYW1tYWQgIiwidW5pcXVlX25hbWUiOiJtZGFueWFsQG1lZGNhcmVtc28uY29tIiwibmFtZWlkIjoiMTAyODQwIiwiUHJhY3RpY2VDb2RlIjoiMCIsIlByb3ZpZGVyQ29kZSI6IjAiLCJVc2VyVHlwZSI6IkJpbGxpbmcgZGVwYXJ0bWVudCIsIlBhdGllbnRBY2NvdW50IjoiMCIsIkFwcGxpY2F0aW9uX0lkIjoiMTEwMDIiLCJuYmYiOjE3NzQ4NTA2MzksImV4cCI6MTc3NDkzNzAzOSwiaWF0IjoxNzc0ODUwNjM5LCJpc3MiOiJodHRwOi8vc2VjdXJlbG9naW4ubWVkY2FyZW1zby5jb20vIiwiYXVkIjoiSG4vTUhNVFZJYnAxQWNwWUpZZFRub3VlUkFOUWpBSTNvY1lJY2lidjc1RTVzOFRmYjVTYjZGWG5yNlIrSWJVSlhjdXc0c29PK1NPYnVGWm5HMUoraEhsQWQwWmw4c081eTJNcHdTVUNoVEk9In0.D4scBzuGHAjQt3gP5EibkX7C-4ZakXIIwC-mBwCYZhE";
	//staging https://staging-maxapi.medcaremso.com
	//dev https://betaqa-maxapi.medcaremso.com
	//live  https://maxapi.medcaremso.com
	String baseURL="https://maxapi.medcaremso.com";
	static String practiceCode;
	
	static String sheetName = "Sheet1";
	static String uploadURL="https://doc.medcaremso.com/api/v1/filerecever/uploadsinglefile";
	
	
	
	public boolean checkPayment(String claimNum, String patientAccountNum ) throws IOException, InterruptedException {
		
		
		
		HttpRequest request1 = HttpRequest.newBuilder()
				.uri(URI.create(baseURL+"/api/Claim/GetClaimPayments?ClaimNo="+claimNum+"&PatientAccount="+patientAccountNum+"&PracticeCode="+practiceCode+"&isRectifiedPayment=false&allRows=false"))
				.header("accept", "application/json")
				.header("accept-language", "en-US,en;q=0.9")
				.header("access-control-allow-credentials", "true")
				.header("access-control-allow-headers", "*")
				.header("access-control-allow-methods", "*")
				.header("access-control-allow-origin", "*")
				.header("authorization", "Bearer "+token+"")
				.header("content-type", "application/json")
				.header("practicecode", practiceCode)
				.header("sec-ch-ua-mobile", "?0")
				.header("sec-fetch-dest", "empty")
				.header("sec-fetch-mode", "cors")
				.header("sec-fetch-site", "same-site")
				.GET()				
				.build();
		HttpResponse<String> response1 = null;
		response1  = HttpClient.newHttpClient().send(request1, HttpResponse.BodyHandlers.ofString());
		
		ObjectMapper objectMapper = new ObjectMapper();
        JsonNode rootNode = objectMapper.readTree(response1.body());	  
        System.out.println(     rootNode.path("status").asText());
		if(!rootNode.path("status").asText().equals("false")) {
		
        JsonNode paymentInfo = rootNode.path("payload");
    //	System.out.println("Payment Exists");
    	return true;
		}else {
		//	System.out.println("Payment does not exist");
			return false;
		}
	}
	

	
	public String getPatientAccNum(String claimNum) throws IOException, InterruptedException {
		HttpRequest request1 = HttpRequest.newBuilder()
				.uri(URI.create(baseURL+"/api/Claim/GetClaim?claimNo="+claimNum+""))
				.header("accept", "application/json")
				.header("accept-language", "en-US,en;q=0.9")
				.header("access-control-allow-credentials", "true")
				.header("access-control-allow-headers", "*")
				.header("access-control-allow-methods", "*")
				.header("access-control-allow-origin", "*")
				.header("authorization", "Bearer "+token+"")
				.header("content-type", "application/json")
				.header("practicecode", practiceCode)
				.header("sec-ch-ua-mobile", "?0")
				.header("sec-fetch-dest", "empty")
				.header("sec-fetch-mode", "cors")
				.header("sec-fetch-site", "same-site")
				.GET()				
				.build();
		HttpResponse<String> response1 = null;
		response1  = HttpClient.newHttpClient().send(request1, HttpResponse.BodyHandlers.ofString());
		System.out.println( response1.statusCode());
		ObjectMapper objectMapper = new ObjectMapper();
        JsonNode rootNode = objectMapper.readTree(response1.body());
        JsonNode payload = rootNode.path("payload");
        System.out.println(   payload.path("Patient_Account").asText());
		return  payload.path("Patient_Account").asText();
	}
	
	
	
	
	 public static String uploadFile(String filePath) throws IOException {
	        File file = new File(filePath);

	        // Check if file exists
	        if (!file.exists()) {
	            System.out.println("File not found at: " + filePath);
	            return "Not Found";
	        }
	
	        OkHttpClient client = new OkHttpClient();

	        RequestBody fileBody = RequestBody.create(file, MediaType.parse("application/pdf"));

	        MultipartBody requestBody = new MultipartBody.Builder()
	                .setType(MultipartBody.FORM)
	                .addFormDataPart("file", file.getName(), fileBody)  // Ensure "file" matches API parameter
	                .build();

	        Request request = new Request.Builder()
	                .url(uploadURL)
	                .post(requestBody)
	                .addHeader("servertoken", "Maximus_SASKKKCXKXNJSJW)@@WMSLLSCMLJSSASNXOSO)WEJSNCKBI@#@@@##") // Remove if not required
	                .addHeader("clientid", practiceCode)
	                .addHeader("folderclient", "BatchFile")
	                .addHeader("subclientid", practiceCode)
	                .build();
	        
	        Response response = client.newCall(request).execute();
	        JsonNode  fileUrl = null;
	        if (response.isSuccessful()) {
                // Read the response body (assumes JSON response)
                String responseBody = response.body().string();
                System.out.println("Response Body: " + responseBody);
                ObjectMapper mapper = new ObjectMapper();
               
             
                    JsonNode jsonResponse = mapper.readTree(responseBody);
                    JsonNode payload = jsonResponse.path("payload");
                    
                    if (payload.has("fileUrl")) {
                    	  fileUrl = payload.path("fileUrl");
                    	  return fileUrl.asText();
                    }
		               
	        }else if(!response.isSuccessful()) {
	        	return "No Response";
	    }
		
	        return null;
	        }
	 
	 public String createBatch(String amount, String userAssigned, String filePathMaximus, String fileName,String batchDate)  throws IOException, InterruptedException {
		 
		 String jsonBody= "{\"Payment_Source\":\"Web Portal\",\"Batch_Amount\":\""+amount+"\",\"No_of_Checks\":\"1\",\"Assigned_To_User\":\""+userAssigned+"\",\"Assigned_To\":\""+userAssigned+"\",\"File_Path\":\""+filePathMaximus+"\",\"File_Name\":\""+fileName+"\",\"No_of_Pages\":\"1\",\"Practice_Code\":\""+practiceCode+"\",\"Batch_Status\":\"Pending\",\"Batch_Date\":\""+batchDate+"\"}";

		 System.out.println(jsonBody);
		 
		 HttpRequest request = HttpRequest.newBuilder()
				 .uri(URI.create(baseURL+"/api/Payment/SavePaymentBatch"))
					.header("accept", "application/json")
					.header("accept-language", "en-US,en;q=0.9")
					.header("access-control-allow-credentials", "true")
					.header("access-control-allow-headers", "*")
					.header("access-control-allow-methods", "*")
					.header("access-control-allow-origin", "*")
					.header("authorization", "Bearer "+token+"")
					.header("content-type", "application/json")
					.header("practicecode", practiceCode)
					.header("sec-ch-ua-mobile", "?0")
					.header("sec-fetch-dest", "empty")
					.header("sec-fetch-mode", "cors")
					.header("sec-fetch-site", "same-site")
					.method("POST",HttpRequest.BodyPublishers.ofString(jsonBody))				
					.build();
			HttpResponse<String> response = null;
		
			
			
				response  = HttpClient.newHttpClient().send(request, HttpResponse.BodyHandlers.ofString());

				
				ObjectMapper objectMapper = new ObjectMapper();
	            JsonNode rootNode = objectMapper.readTree(response.body());	  
	            
	            JsonNode address = rootNode.path("payload");
	            System.out.println(response.body());
	            
	            System.out.println(address.get("Batch_Payment_Id").asText());
			
		return address.get("Batch_Payment_Id").asText();
		 
	 }
	 
	 
	
	public Map<String, String> getNPIandStateofPractice() {
		 
		String jsonBody= "{\"pageIndex\":1,\"pageSize\":10,\"RowOfPage\":1,\"PageNumber\":10,\"practiceCode\":\""+practiceCode+"\",\"searchText\":\"\"}";
		
		HttpRequest request1 = HttpRequest.newBuilder()
				.uri(URI.create(baseURL+"/api/Practice/GetPracticeProviderList"))
				.header("accept", "application/json")
				.header("accept-language", "en-US,en;q=0.9")
				.header("access-control-allow-credentials", "true")
				.header("access-control-allow-headers", "*")
				.header("access-control-allow-methods", "*")
				.header("access-control-allow-origin", "*")
				.header("authorization", "Bearer "+token+"")
				.header("content-type", "application/json")
				.header("practicecode", practiceCode)
				.header("sec-ch-ua-mobile", "?0")
				.header("sec-fetch-dest", "empty")
				.header("sec-fetch-mode", "cors")
				.header("sec-fetch-site", "same-site")
				.method("POST",HttpRequest.BodyPublishers.ofString(jsonBody))				
				.build();
		HttpResponse<String> response1 = null;
		Map<String, String> npiAddressMap = new HashMap<>();
		try {
		
			response1  = HttpClient.newHttpClient().send(request1, HttpResponse.BodyHandlers.ofString());

			
			ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(response1.body());	  
            
            JsonNode address = rootNode.path("payload");
            for (int i=0; i<address.size(); i++) {
            	
            	if(address.get(i).path("Is_Active").asText().equals("true")) {
            		npiAddressMap.put(address.get(i).path("Full_Name").asText(), address.get(i).path("NPI").asText());
            	}
            	
            	
            	
            }
            npiAddressMap.put("State", address.get(0).path("Address").asText());
          
          
           
 }catch(Exception e) {}
		return npiAddressMap;
 }
	
	
	public int saveNote(String claimNum, String notesCategoryID, String CPTS, String allowedLineAmount,String statusCode, String Description,String practiceCode) throws IOException, InterruptedException {
		 
		
		
		
		String jsonBody= "{\"notes_Category_Id\":"+notesCategoryID+",\"description\":\"The claim has issue with the following details:\\nCPTs= "+CPTS+"\\nLine Allowed = "+allowedLineAmount+"\\nStatus Codes= "+statusCode+"\\nDescription = "+Description+"\\nEOBs placed at the path \\\\\\\\10.170.193.71\\\\Development Team\\\\RPA\\\\ECW\\\\Availity Followups\\\\DownloadedFiles\",\"claim_No\":"+claimNum+"}";
		
		HttpRequest request1 = HttpRequest.newBuilder()
				.uri(URI.create(baseURL+"/api/ClaimNotes/SaveClaimNotes"))
				.header("accept", "application/json")
				.header("accept-language", "en-US,en;q=0.9")
				.header("access-control-allow-credentials", "true")
				.header("access-control-allow-headers", "*")
				.header("access-control-allow-methods", "*")
				.header("access-control-allow-origin", "*")
				.header("authorization", "Bearer "+token+"")
				.header("content-type", "application/json")
				.header("practicecode", practiceCode)
				.header("sec-ch-ua-mobile", "?0")
				.header("sec-fetch-dest", "empty")
				.header("sec-fetch-mode", "cors")
				.header("sec-fetch-site", "same-site")
				.method("POST",HttpRequest.BodyPublishers.ofString(jsonBody))				
				.build();
		HttpResponse<String> response1 = null;

		int statusCodeOfResponse = 0;
		try {
		
			response1  = HttpClient.newHttpClient().send(request1, HttpResponse.BodyHandlers.ofString());
			statusCodeOfResponse = response1.statusCode();
			
			ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(response1.body());	  
            
            JsonNode payload = rootNode.path("payload");
        
         
          
          
           
 }catch(Exception e) {}
		return statusCodeOfResponse;
 }                                                                                                                             
	
	
	public int updateClaimCategory(String claimNum,String practiceCode,String claimID) {
		 
		String jsonBody= "{\"Claim_Nos\":["+claimNum+"],\"Claim_Category_ID\":"+claimID+"}" ;
		
		HttpRequest request1 = HttpRequest.newBuilder()
				.uri(URI.create(baseURL+"/api/OpenBucket/AssignCategoryToClaims"))
				.header("accept", "application/json")
				.header("accept-language", "en-US,en;q=0.9")
				.header("access-control-allow-credentials", "true")
				.header("access-control-allow-headers", "*")
				.header("access-control-allow-methods", "*")
				.header("access-control-allow-origin", "*")
				.header("authorization", "Bearer "+token+"")
				.header("content-type", "application/json")
				.header("practicecode", practiceCode)
				.header("sec-ch-ua-mobile", "?0")
				.header("sec-fetch-dest", "empty")
				.header("sec-fetch-mode", "cors")
				.header("sec-fetch-site", "same-site")
				.method("POST",HttpRequest.BodyPublishers.ofString(jsonBody))				
				.build();
		HttpResponse<String> response1 = null;
		Map<String, String> npiAddressMap = new HashMap<>();
		int statusCodeOfResponse = 0;
		try {
		
			response1  = HttpClient.newHttpClient().send(request1, HttpResponse.BodyHandlers.ofString());
			statusCodeOfResponse = response1.statusCode();
			
			ObjectMapper objectMapper = new ObjectMapper();
            JsonNode rootNode = objectMapper.readTree(response1.body());	  
            
            JsonNode payload = rootNode.path("payload");
        
         
          
          
           
 }catch(Exception e) {}
		return statusCodeOfResponse;
 }
	
	
	
	 public void downloadExcel() {
			
	        String jsonBody = "{\r\n    \"PartialFilterOnGroup\": false,\r\n    \"Status\": \"New\",\r\n    \"ArFilterOnObjDetails\": {\r\n        \"ClaimTypeDetail\": \"all\",\r\n        \"IsPartialClaimFilter\": false,\r\n        \"ClaimCategory\": \"all\",\r\n        \"AssignedTo\": \"all\",\r\n        \"ClaimNoFltOnDetail\": null,\r\n        \"PageNumber\": 1,\r\n        \"RowOfPage\": 10\r\n    },\r\n    \"TabIndex\": 0,\r\n    \"payerCheckbox\": true,\r\n    \"agingCheckbox\": true,\r\n    \"ClaimNoFltOnGroup\": null,\r\n    \"LocationFilterOnGroup\": \"\",\r\n    \"DenialFilterOnGroup\": \"\",\r\n    \"AgingFilterOnGroup\": \"0-30,31-60,61-90,91-120,+120\",\r\n    \"ProvideFilterOnGroup\": \"\",\r\n    \"PayerFilterOnGroup\": \"\",\r\n    \"PRACTICE_CODE\": \""+practiceCode+"\",\r\n    \"GroupType\": \"Payer-No-Responses\",\r\n    \"RightWiseData\": \"New,In-Progress,Completed,\",\r\n    \"PayerListOnGroup\": [],\r\n    \"DenialListOnGroup\": [],\r\n    \"AgingListOnGroup\": [\r\n        {\r\n            \"id\": 2,\r\n            \"name\": \"31-60\"\r\n        },\r\n        {\r\n            \"id\": 3,\r\n            \"name\": \"61-90\"\r\n        },\r\n        {\r\n            \"id\": 4,\r\n            \"name\": \"91-120\"\r\n        },\r\n        {\r\n            \"id\": 5,\r\n            \"name\": \"+120\"\r\n        }\r\n    ],\r\n    \"ProvideListOnGroup\": [],\r\n    \"LocationListOnGroup\": [],\r\n    \"GroupTypeValue\": \"Total_\",\r\n    \"AllClaimsFilter\": true\r\n}";

	        // Create HttpClient
	        HttpClient client = HttpClient.newHttpClient();

	        // Create HttpRequest
	        HttpRequest request = HttpRequest.newBuilder()
	                .uri(URI.create(baseURL+"/api/Claim/GenerateCollectionExcel"))
	                .header("accept", "application/json")
	                .header("accept-language", "en-US,en;q=0.9")
	                .header("access-control-allow-credentials", "true")
	                .header("access-control-allow-headers", "*")
	                .header("access-control-allow-methods", "*")
	                .header("access-control-allow-origin", "*")
	                .header("authorization", "Bearer "+token+"")
	                .header("content-type", "application/json")
	                .header("practicecode", practiceCode)
	                .header("priority", "u=1, i")
	                .header("sec-ch-ua", "\"Not)A;Brand\";v=\"99\", \"Google Chrome\";v=\"127\", \"Chromium\";v=\"127\"")
	                .header("sec-ch-ua-mobile", "?0")
	                .header("sec-ch-ua-platform", "\"Windows\"")
	                .header("sec-fetch-dest", "empty")
	                .header("sec-fetch-mode", "cors")
	                .header("sec-fetch-site", "same-site")
	                .POST(HttpRequest.BodyPublishers.ofString(jsonBody))
	                .build();

	        try {
	            // Send request and receive response
	            HttpResponse<InputStream> response = client.send(request, HttpResponse.BodyHandlers.ofInputStream());

	            // Check response status code
	            if (response.statusCode() == 200) {
	                // Save the response input stream to a file
	                try (InputStream in = response.body();
	                     FileOutputStream fileOutputStream = new FileOutputStream("Availity Report "+practiceCode+".xlsx")) {

	                    byte[] buffer = new byte[1024];
	                    int bytesRead;

	                    while ((bytesRead = in.read(buffer)) != -1) {
	                        fileOutputStream.write(buffer, 0, bytesRead);
	                    }

	                    System.out.println("File downloaded successfully!");
	                }
	            } else {
	                // Print error message
	                System.err.println("Failed to download file. HTTP Status Code: " + response.statusCode());
	                try (InputStream errorStream = response.body()) {
	                    String errorResponse = new String(errorStream.readAllBytes());
	                    System.err.println("Error Response: " + errorResponse);
	                } catch (IOException e) {
	                    System.err.println("Error reading error response: " + e.getMessage());
	                }
	            }
	        } catch (IOException | InterruptedException e) {
	            e.printStackTrace();
	            System.err.println("Error sending request: " + e.getMessage());
	        }
	    }
		 
	
	
	 
	 public static String extractStateAcronym(String address) {
	        // Regular expression to match the state acronym (2 uppercase letters)
	        String regex = ",\\s*([A-Z]{2}),";
	        Pattern pattern = Pattern.compile(regex);
	        Matcher matcher = pattern.matcher(address);

	        // Find the last occurrence of the state acronym
	        String stateAcronym = null;
	        while (matcher.find()) {
	            stateAcronym = matcher.group(1); // Update to the latest match
	        }
	        
	        return stateAcronym; // Return the last found state acronym or null if none found
	    }
	  public static String vlookup(Worksheet worksheet, String lookupValue, int lookupColumn, int resultColumn) {
	        // Iterate through each row in the worksheet
	        for (int rowIndex = 1; rowIndex < worksheet.getCells().getMaxDataRow() + 1; rowIndex++) {
	            // Get the value in the lookup column for the current row
	            Cell cell = worksheet.getCells().get(rowIndex, lookupColumn);
	            if (cell.getType() == CellValueType.IS_STRING) {
	                String value = cell.getStringValue();
	                // Check if values match for the lookup column
	                if (lookupValue.toLowerCase().equals(value.toLowerCase())) {
	                    // Get the corresponding value from the result column
	                    return worksheet.getCells().get(rowIndex, resultColumn).getStringValue();
	                }
	            }
	        }

	        // If no match is found, return an empty string or whatever is appropriate for your use case
	        return "No match found";
	    }
	    
	    public static void performVLookup(String excel1, String excel2) throws Exception {
      Workbook workbook1 = new Workbook(System.getProperty("user.dir")+"\\"+excel1);
      Workbook workbook2 = new Workbook(System.getProperty("user.dir") + "\\"+excel2);
     
      Worksheet worksheet1 = workbook1.getWorksheets().get(0);
      Worksheet worksheet2 = workbook2.getWorksheets().get(0);

      // Get the data range from Excel1 (1 column to lookup)
      int lookupColumnLIS = 9; // 0-based index of the column to lookup in  Excel2 (LIS Master)

      // Get the column indices in Excel2 where to place the VLOOKUP results
      int targetColumn1 = 0; //Claim Number // 0-based index of the first column in Excel1 to place the VLOOKUP result
     
      
      // Iterate through each row in Excel2
      for (int rowIndex = 1; rowIndex < worksheet1.getCells().getMaxDataRow() + 1; rowIndex++) {
          // Get the value to lookup from Excel2
          Cell cell = worksheet1.getCells().get(rowIndex, lookupColumnLIS);
          String lookupValue = cell.getStringValue();

          // Perform VLOOKUP in Excel1 for the first target column
          String resultValue1 = vlookup(worksheet2, lookupValue, 0, 2);
          String resultValueState = vlookup(worksheet2, lookupValue, 0, 1);
     
//if(resultValue1.toLowerCase().contains("do not process")) {
//	System.out.println("here");
//}
          // Combine the results and place in Excel2
          if(resultValue1.toLowerCase().contains("do not process")) {
        	  worksheet1.getCells().get(rowIndex, 9).putValue("do not process for availity");
        	  worksheet1.getCells().get(rowIndex, 48).putValue("do not process for availity");
        	  worksheet1.getCells().get(rowIndex, 46).putValue("do not process for availity");
          }else {
          worksheet1.getCells().get(rowIndex, 9).putValue(resultValue1);
         
          worksheet1.getCells().get(rowIndex, 48).putValue(resultValueState);
          }
          
          if(resultValue1.equals("No match found")) {
        	  worksheet1.getCells().get(rowIndex, 46).putValue(resultValue1+" for payer");
        	  
          }
        
      }
 

      // Save the modified Excel2
      workbook1.save(excel1);
	    }
	    
	    
	    public static void divideExcel(String inputFile, String practiceCode) throws IOException {
	        try (FileInputStream fis = new FileInputStream(inputFile);
	             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

	            Sheet original = workbook.getSheetAt(0);
	            Row header = original.getRow(0);
	            if (header == null) return;

	            DataFormatter fmt = new DataFormatter();

	            // Collect indices of rows to keep (exclude "do not process for availity" in col index 9)
	            List<Integer> valid = new ArrayList<>();
	            int lastRow = original.getLastRowNum();
	            for (int r = 1; r <= lastRow; r++) { // start after header
	                Row row = original.getRow(r);
	                if (row == null) continue;
	                org.apache.poi.ss.usermodel.Cell c = row.getCell(9) ; // 10th column
	                String v = fmt.formatCellValue(c).trim().toLowerCase(Locale.ROOT);
	                if (!"do not process for availity".equals(v) && !v.contains("no match found")) {
	                    valid.add(r);
	                }
	            }

	            // Split valid rows into 5 parts as evenly as possible
	            int parts = 5;
	            int base = valid.size() / parts;
	            int rem  = valid.size() % parts; // first 'rem' parts get one extra

	            int offset = 0;
	            for (int i = 0; i < parts; i++) {
	                int size = base + (i < rem ? 1 : 0);

	                try (XSSFWorkbook outWb = new XSSFWorkbook()) {
	                    Sheet out = outWb.createSheet("Sheet1");

	                    // copy header
	                    copyRow(header, out.createRow(0));

	                    // copy rows for this chunk
	                    for (int k = 0; k < size; k++) {
	                        int srcRowIdx = valid.get(offset + k);
	                        Row src = original.getRow(srcRowIdx);
	                        Row dst = out.createRow(k + 1); // +1 to keep header at row 0
	                        copyRow(src, dst);
	                    }

	                    // write file
	                    try (FileOutputStream fos = new FileOutputStream("Availity " + (i + 1) + " "+practiceCode+".xlsx")) {
	                        outWb.write(fos);
	                    }
	                }

	                offset += size;
	            }
	        }
	    }

	    
	    
	    
	   /* 
	    
	    
	   public static void divideExcel(String inputFile) throws IOException {
	        FileInputStream fis = new FileInputStream(inputFile);
	        XSSFWorkbook workbook = new XSSFWorkbook(fis);
	        Sheet originalSheet = workbook.getSheetAt(0);

	        int totalRows = originalSheet.getPhysicalNumberOfRows();
	     //   int rowsPerPart = totalRows / 5;
	        
	        int validRows = 0;

	        // Count valid rows (excluding "do not process" in column 5)
	        for (int rowIndex = 1; rowIndex < totalRows; rowIndex++) { // Start from 1 to skip the header
	            Row row = originalSheet.getRow(rowIndex);
	            if (row != null && row.getCell(9) != null) {  // Check if 5th column exists and is not null
	                String cellValue = row.getCell(9).getStringCellValue();
	                if (!"do not process for availity".equalsIgnoreCase(cellValue)) {
	                    validRows++; // Count valid rows
	                }
	            }
	        }

	        // Calculate rows per part based on valid rows
	        int rowsPerPart = validRows / 5;
	        
System.out.println(validRows);
	        // Copy header row
	        Row headerRow = originalSheet.getRow(0);

	        for (int i = 0; i < 5; i++) {
	            XSSFWorkbook newWorkbook = new XSSFWorkbook();
	            Sheet newSheet = newWorkbook.createSheet("Sheet1");

	            // Copy the header to the new sheet
	            Row newHeaderRow = newSheet.createRow(0);
	            copyRow(headerRow, newHeaderRow);

	            int startRow = i * rowsPerPart + 1; // Start from the row after the header
	            int endRow = (i == 4) ? validRows : startRow + rowsPerPart;

	            for (int rowIndex = startRow; rowIndex < endRow; rowIndex++) {
	                Row originalRow = originalSheet.getRow(rowIndex);
	              
	                if(!originalRow.getCell(9).getStringCellValue().equalsIgnoreCase("do not process for availity")) {
	                	  Row newRow = newSheet.createRow(rowIndex - startRow + 1); // Adjust for header
	                copyRow(originalRow, newRow);
	                }
	            }

	            FileOutputStream fos = new FileOutputStream("Availity " + (i + 1) + ".xlsx");
	            newWorkbook.write(fos);
	            fos.close();
	            newWorkbook.close();
	        }

	        workbook.close();
	        fis.close();
	    }
*/
	    
	    
	    
	/*    public static void divideExcel(String inputFile) throws IOException {
	        FileInputStream fis = new FileInputStream(inputFile);
	        XSSFWorkbook workbook = new XSSFWorkbook(fis);
	        Sheet originalSheet = workbook.getSheetAt(0);

	        int totalRows = originalSheet.getPhysicalNumberOfRows();
	        int validRows = 0;

	        // Count valid rows (excluding "do not process" in column 5)
	        for (int rowIndex = 1; rowIndex < totalRows; rowIndex++) { // Start from 1 to skip the header
	            Row row = originalSheet.getRow(rowIndex);
	            if (row != null && row.getCell(9) != null) {  // Check if 5th column exists and is not null
	                String cellValue = row.getCell(9).getStringCellValue();
	                if (!"do not process for availity".equalsIgnoreCase(cellValue)) {
	                    validRows++; // Count valid rows
	                }
	            }
	        }

	        // Calculate rows per part based on valid rows
	        int rowsPerPart = validRows / 5;

	        // Copy header row
	        Row headerRow = originalSheet.getRow(0);

	        // This will track the index for valid rows
	        for (int i = 0; i < 5; i++) {
	        	
	            XSSFWorkbook newWorkbook = new XSSFWorkbook();
	            Sheet newSheet = newWorkbook.createSheet("Sheet1");
	            int currentValidRowIndex = 1;
	            // Copy the header to the new sheet
	            Row newHeaderRow = newSheet.createRow(0);
	            copyRow(headerRow, newHeaderRow);

	            int startRow = i * rowsPerPart + 1; // Start from the row after the header
	            int endRow = (i == 4) ? validRows : startRow + rowsPerPart;

	            // Copy valid rows only
	            for (int rowIndex = 1; rowIndex < totalRows; rowIndex++) {
	                Row originalRow = originalSheet.getRow(rowIndex);
	                if (originalRow != null && originalRow.getCell(9) != null) {
	                    String cellValue = originalRow.getCell(9).getStringCellValue();
	                    if (!"do not process for availity".equalsIgnoreCase(cellValue) ) {
	                        if (currentValidRowIndex >= startRow && currentValidRowIndex < endRow) {
	                            Row newRow = newSheet.createRow(currentValidRowIndex - startRow + 1); // Adjust for header
	                            copyRow(originalRow, newRow);
	                            currentValidRowIndex++;
	                        }
	                    }
	                }
	            }

	            // Write the new workbook to a file
	            FileOutputStream fos = new FileOutputStream("Availity " + (i + 1) + ".xlsx");
	            newWorkbook.write(fos);
	            fos.close();
	            newWorkbook.close();
	        }

	        workbook.close();
	        fis.close();
	    }

	 

	    
	 */   
	    
	    private static void copyRow(Row originalRow, Row newRow) {
	        if (originalRow != null) {
	            for (int j = 0; j < originalRow.getLastCellNum(); j++) {
	               org.apache.poi.ss.usermodel.Cell originalCell = originalRow.getCell(j);
	                org.apache.poi.ss.usermodel.Cell newCell = newRow.createCell(j);
	                if (originalCell != null) {
	                    switch (originalCell.getCellType()) {
	                        case STRING:
	                            newCell.setCellValue(originalCell.getStringCellValue());
	                            break;
	                        case NUMERIC:
	                            newCell.setCellValue(originalCell.getNumericCellValue());
	                            break;
	                        case BOOLEAN:
	                            newCell.setCellValue(originalCell.getBooleanCellValue());
	                            break;
	                        case FORMULA:
	                            newCell.setCellFormula(originalCell.getCellFormula());
	                            break;
	                        default:
	                            newCell.setBlank();
	                            break;
	                    }
	                }
	            }
	        }
	    }

	    
	    
	    public static void addHeaders(String inputFile, int startColumn) throws IOException {
	    	
	    	 String[] headers = {
	                 "EOB Downloaded", "Availity DOS", "Claim Number", "Check Number", "Check Date",
	                 "Finalized Date", "Payment Date", "Paid Amount","Billed Amount", "Allowed Amount",
	                 "Received Date", "Denial Reason", "Action", "Line Allowed", "Line CPT", "Line Paid", "Line Billed",
	                 "Line Hippa", "Line Remark Code","Line Status Code","Line Descriptions", "Line Copay", "Line Deductible",
	                 "Line Ineligible", "Line Coinsurance", "Bot Status", "Maximus Status","State"
	             };
	    	
	        FileInputStream fis = new FileInputStream(inputFile);
	        XSSFWorkbook workbook = new XSSFWorkbook(fis);
	        Sheet sheet = workbook.getSheetAt(0); // Modify if you need a different sheet

	        // Create a new row for the headers, or get the first row if it already exists
	        Row headerRow = sheet.getRow(0);
	        if (headerRow == null) {
	            headerRow = sheet.createRow(0);
	        }

	        // Add headers starting from the specified column
	        for (int i = 0; i < headers.length; i++) {
	            org.apache.poi.ss.usermodel.Cell cell = headerRow.createCell(startColumn + i);
	            cell.setCellValue(headers[i]);
	        }

	        // Write changes to a new file or overwrite the original
	        FileOutputStream fos = new FileOutputStream(inputFile);
	        workbook.write(fos);
	        fos.close();
	        workbook.close();
	        fis.close();
	    }
	    
	    
	    
	 
}


	 

