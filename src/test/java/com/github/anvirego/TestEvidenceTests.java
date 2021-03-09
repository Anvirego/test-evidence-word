package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.anvirego.interfaces.TestEvidenceInterface;

public class TestEvidenceTests {
	
	public static void main(String args[]) throws FileNotFoundException, IOException {
		//cap = new Screenshot(driver, TestEvidence.getInstance(base.getPlatform(), scenario.getName(), base.getTesterName(), base.getEnvironment()), scenario.getName(), base.getLanguage());	

		TestEvidenceInterface ti = TestEvidenceInstance.getInstance("Android", "Scenario", "Angelica", "UAT");
		
		ti.setHeight(0);
		ti.setWidth(0);
		ti.createEvidence("path", "description", 0);
		
		ti.closeWord("path", null);
		
		System.out.println("FolderName: "+ti.returnFolderName());
	}//Main

}//Class
