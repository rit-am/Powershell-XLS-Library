// https://stackoverflow.com/questions/29545611/executing-powershell-commands-in-java-program
// https://github.com/profesorfalken/jPowerShell
		  // /src/home/ps/ssps.ps1
		  // /src/home/ShowXLSCellData.java
		  // /src/home/xls/cstmxl.xlsx

package home;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class ShowXLSCellData {

	 public static void main(String[] args) throws IOException {
		 String varstrPS1filename ="..\\ps1\\ssps.ps1";
		 String varstrXLSfilename="..\\xls\\cstmxl.xlsx";
		 String varstrXLSsheetname="CustomSheet";
		 Integer varintROW=2;
		 Integer varintCOL=9;
		 String command = "powershell.exe  "+varstrPS1filename+
				  " -file "+ varstrXLSfilename +
				  " -sheet "+ varstrXLSsheetname + 
				  " -row " + varintROW + 
				  " -col " + varintCOL;
		  System.out.println("Command :: "+command);
		  // Executing the command
		  Process powerShellProcess = Runtime.getRuntime().exec(command);
		  // Getting the results
		  powerShellProcess.getOutputStream().close();
		  String line;
		  System.out.println("Standard Output:");
		  BufferedReader stdout = new BufferedReader(new InputStreamReader(
		    powerShellProcess.getInputStream()));
		  while ((line = stdout.readLine()) != null) {
		   System.out.println(line);
		  }
		  stdout.close();
		  System.out.println("Standard Error:");
		  BufferedReader stderr = new BufferedReader(new InputStreamReader(
		    powerShellProcess.getErrorStream()));
		  while ((line = stderr.readLine()) != null) {
		   System.out.println(line);
		  }
		  stderr.close();
		  System.out.println("Done");

		 }


}



