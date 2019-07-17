package exceldata;

import java.io.IOException;
import java.util.ArrayList;

public class getexcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		exceldata d = new exceldata();
         
		ArrayList<String> data = d.getdata("contact");
		
		
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
	}

}
