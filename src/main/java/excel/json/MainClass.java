package excel.json;

public class MainClass {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String url = "https://raw.githubusercontent.com/OAI/OpenAPI-Specification/main/examples/v3.0/petstore.json";
		JsonToExcel.exportswaggerUrlToExcel(url);
		
	}

}
