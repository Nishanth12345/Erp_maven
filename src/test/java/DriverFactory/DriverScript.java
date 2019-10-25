package DriverFactory;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunlib.FunctionLibrary;
import Utilities.ExcelFileUtil;

public class DriverScript {
	WebDriver driver;
	ExtentReports report;
	ExtentTest test;
	public void startTest() throws Throwable
	{
		//creating reference object foe excel util methods
		ExcelFileUtil excel=new ExcelFileUtil();
		//iterating alll row in MasterTestCases sheet
		for(int i=1;i<=excel.rowCount("MasterTestCases");i++)
		{
			String ModuleStatus="";
			if(excel.getData("MasterTestCases",i,2).equalsIgnoreCase("Y"))
			{
				//store module name into TCModule
				String TCModule=excel.getData("MasterTestCases", i, 1);
				report=new ExtentReports("./Reports/"+TCModule+FunctionLibrary.generateDate()+".html");
				//iterate all rows in TCModule Sheet
				for(int j=1;j<=excel.rowCount(TCModule);j++)
				{
					test=report.startTest(TCModule);
					//read all columns in TCModule testcase
					String Description=excel.getData(TCModule, j, 0);
					String Object_Type=excel.getData(TCModule, j, 1);
					String Locator_Type=excel.getData(TCModule, j, 2);
					String Locator_Value=excel.getData(TCModule, j, 3);
					String Test_Data=excel.getData(TCModule, j, 4);
					//calling methods from function library
					try
					{
						if(Object_Type.equalsIgnoreCase("startBrowser"))
						{
							driver=FunctionLibrary.startBrowser(driver);
							test.log(LogStatus.INFO,Description);
							System.out.println("Executing startbrowser");
						}
						else if(Object_Type.equalsIgnoreCase("openApplication"))
						{
							FunctionLibrary.openApplication(driver);
							test.log(LogStatus.INFO, Description);
							System.out.println("Executing openApplication");
						}
						else if(Object_Type.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO,Description);
							System.out.println("Executing waitForElement");
						}else if(Object_Type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
							System.out.println("Executing typeAction");
						}
						else if(Object_Type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
							System.out.println("Executing clickAction");
						}
						else if(Object_Type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser(driver);
							test.log(LogStatus.INFO, Description);
							System.out.println("Executing closeBrowser");
						}
						
						else if(Object_Type.equalsIgnoreCase("captureData"))
						{
							FunctionLibrary.captureData(driver,Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
							System.out.println("Executing capturedata");
						}
						else if(Object_Type.equalsIgnoreCase("tableValidation"))
						{
							FunctionLibrary.tableValidation(driver, Test_Data);
							test.log(LogStatus.INFO, Description);
						}
						else if(Object_Type.equalsIgnoreCase("stockCategories"))
						{
							
							FunctionLibrary.stockCategories(driver);
							test.log(LogStatus.INFO,Description);
							
						}
						else if(Object_Type.equalsIgnoreCase("tableValidation1"))
						{
							FunctionLibrary.tableValidation1(driver, Test_Data);
							test.log(LogStatus.INFO,Description);
						}
							
						//write as pass into status column
						excel.setCellData(TCModule, j, 5, "PASS");
						ModuleStatus="true";
					}catch(Exception e)
					{
						excel.setCellData(TCModule, j, 5, "FAIL");
						ModuleStatus="false";
						System.out.println(e.getMessage());
						break;
					}
					if(ModuleStatus.equalsIgnoreCase("TRUE"))
					{
						excel.setCellData("MasterTestCases", i, 3, "pass");
					}
					else if(ModuleStatus.equalsIgnoreCase("FALSE"))
					{
						excel.setCellData("MasterTestCases", i, 3, "Fail");
					}
				}
				report.flush();
				report.endTest(test);
			}
		
			else
			{
				//write as not executed in status column for flag N
				excel.setCellData("MasterTestCases", i, 3, "Not Executed");
				}
			}
		}
	}
