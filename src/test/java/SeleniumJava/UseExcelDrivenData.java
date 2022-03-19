package SeleniumJava;

import java.io.IOException;
import java.util.ArrayList;

public class UseExcelDrivenData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		demoExcelDriven d = new demoExcelDriven();
		ArrayList al = d.excelDriven("S.No.4");
		System.out.println(al.get(0));
		System.out.println(al.get(1));
		System.out.println(al.get(2));
		System.out.println(al.get(3));
	}

}
