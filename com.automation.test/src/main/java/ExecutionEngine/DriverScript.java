package ExecutionEngine;

import java.lang.reflect.Method;
import java.sql.Driver;

import Utility.ExcelUtilities;
import config.ActionKeyword;

public class DriverScript {

	public static ActionKeyword ac;
	public static String sActionKeyword;
	public static Method mtd[];
	static int iRow;
	public static String sPath;
	// public static Boolean bresult;
	public static boolean flag;

	public DriverScript() throws NoSuchMethodException, SecurityException {
		ac = new ActionKeyword();
		mtd = ac.getClass().getMethods();
	}

	public static void main(String[] args) throws Exception {

		DriverScript hj = new DriverScript();
		sPath = "D:\\DataEngine.xlsx";

		// Adding path file
		ExcelUtilities.setExcelFile(sPath, "ABC");

		for (iRow = 1; iRow <= 4; iRow++) {
			flag = true;
			// This to get the value of column Action Keyword from the excel
			sActionKeyword = ExcelUtilities.getCellData(iRow, 3);

			execute_Actions();
		}
	}

	private static void execute_Actions() throws Exception {

		ac = new ActionKeyword();
		mtd = ac.getClass().getMethods();

		for (int i = 0; i < mtd.length; i++) {

			if (mtd[i].getName().equals(sActionKeyword)) {

				mtd[i].invoke(ac);

				if (flag == true) {
					// If the executed test step value is true, Pass the test step in Excel sheet
					ExcelUtilities.setCellData(ac.Resultp, iRow, 4, "ABC");
					break;
				} 
				else
				{
					// If the executed test step value is false, Fail the test step in Excel sheet
					ExcelUtilities.setCellData(ac.Resultf, iRow, 4, "ABC");
					break;
				}

			}
		}

	}
}
