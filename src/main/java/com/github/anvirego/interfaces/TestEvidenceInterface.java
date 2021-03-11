package com.github.anvirego.interfaces;

import java.io.FileOutputStream;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 1.0 03/2021.
 * TestEvidenceInterface: Interface Implementation. 
 */
public interface TestEvidenceInterface {
	//Method to create the Evidence.
	public void createEvidence(String path, String txtDescription, int step);
	
	//Method to close all instances involved with TestEvidence.
	public Boolean closeWord(String path, FileOutputStream outPicture);
	
	//Sets Picture's heihght.
	public void setHeight(int height);
	
	//Sets Pictures's width.
	public void setWidth(int width);
	
	//Returns folder name.
	public String returnFolderName();

}//Interface