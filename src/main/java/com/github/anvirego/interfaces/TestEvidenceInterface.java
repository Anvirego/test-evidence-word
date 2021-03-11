package com.github.anvirego.interfaces;

import java.io.FileOutputStream;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 1.0 03/2021.
 * TestEvidenceInterface: Interface Implementation. 
 */
public interface TestEvidenceInterface {
	//Method to create the Evidence
	public void createEvidence(String path, String txtDescription, int step);
	
	//
	public Boolean closeWord(String path, FileOutputStream outPicture);
	
	public void setHeight(int height);
	
	public void setWidth(int width);
	
	public String returnFolderName();


}//Interface
