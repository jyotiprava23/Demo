package demo;

public class TestDemo {

	public static void main(String[] args) throws Exception 
  { 
	try
	 {
		System.out.println("Hi");
		ExcelAccess.setExcelFile("D://automation script//LoginDetails.xlsx","Sheet1");
		ExcelAccess.getCellData(1,2);
		ExcelAccess.setCellData("D://automation script//LoginDetails.xlsx","jyoti",1,3); 
	 }
	catch (Exception e)
	 {

		throw (e);
	 }
  }
}


