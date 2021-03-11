package com.github.anvirego;

import java.io.FileNotFoundException;
import java.io.IOException;

import com.github.anvirego.interfaces.TestEvidenceInterface;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 2.0 03/2021. 
 * TestEvidencesTests: Examples of implementation.
 */
public class TestEvidenceTests {
	
	public static void main(String args[]) throws FileNotFoundException, IOException {
		TestEvidenceInterface ti = TestEvidenceInstance.getInstance("Android", "Scenario", "Angelica", "UAT");
		
		ti.setHeight(0);
		ti.setWidth(0);
		ti.createEvidence("path", "description", 0);
		
		ti.closeWord("path", null);
		
		System.out.println("FolderName: "+ti.returnFolderName());
	}//Main

}//Class
