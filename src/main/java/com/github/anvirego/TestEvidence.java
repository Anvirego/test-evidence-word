package com.github.anvirego;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author Ing. Angélica Viridiana Rebolloza Gonzalez.
 * @version 3.1 03/2021.
 * TestEvidence: Creates the WordDocument from scratch.
 */
public final class TestEvidence extends TestEvidenceLogic {
	private static TestEvidence testE;
	private static XWPFDocument wordDocument;
	private int height, width;

	public TestEvidence(String platform, String suiteTestName, String testerName, String environment) {
		//Image size
		height = platform.equalsIgnoreCase("Web") ? 265 : 275;
		width = platform.equalsIgnoreCase("Web") ? 468 : 190;
		wordDocument = createWordDocument(platform, suiteTestName, testerName, environment);
		//folderName = getFolderName();
	}//Constructor
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected static TestEvidence getInstance(String platform, String suiteTestName, String testerName, String environment) {
		System.out.println("==== Get TestEvidence Instance =====");
		if(testE == null) {
			System.out.println("New Instance");
			testE = new TestEvidence(platform, suiteTestName, testerName, environment);
			return testE;
		} else {
			System.out.println("Old Instance");
			return testE;
		}
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public void createEvidence(String path, String txtDescription, int step) {
		System.out.println("::::: createEvidence :::::");
		try {
			XWPFParagraph evidenceP1 = wordDocument.createParagraph();
			evidenceP1.setBorderBottom(Borders.BALLOONS_3_COLORS);
			XWPFRun evidenceR1 = evidenceP1.createRun();
			evidenceR1.setBold(true);
			//Checks if odd number
			if (step % 2 == 0 && step!=1) {
				evidenceR1.addBreak(BreakType.PAGE);
			}
			evidenceR1.setText("Step "+step+": "+txtDescription);
			XWPFParagraph evidenceP2 = wordDocument.createParagraph();
			evidenceP2.setAlignment(ParagraphAlignment.CENTER);
			InputStream picture = new FileInputStream(path+"/.png");
			evidenceR1 = evidenceP2.createRun();
			evidenceR1.addPicture(picture, XWPFDocument.PICTURE_TYPE_PNG, null, Units.toEMU(width), Units.toEMU(height));//height width
			picture.close();
			evidenceR1.addBreak();
		} catch (Exception e) {System.out.println("▓▓▓▓▓▓▓▓▓▓ createEvidence: "+e+" ▓▓▓▓▓▓▓▓▓▓");}
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	public Boolean closeWord(String path, FileOutputStream outPicture) {
		System.out.println("::::: closeWord :::::");
		try {
			if(signEvidence()) {
				wordDocument.write(outPicture);
				outPicture.close();
				System.out.println("\n///Word Document created successfuly///\n");
				testE = null;
				//Deletes directory temporal
				File temporalFiles = new File (path+"/.png");
				temporalFiles = new File(path);
				deleteDirCMD(temporalFiles.getAbsolutePath());
			}
		}catch (org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException r) {
		}catch (Exception e) {System.out.println("▓▓▓▓▓▓▓▓▓▓ closeWord: "+e+" ▓▓▓▓▓▓▓▓▓▓");}
		return true;
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	public void setHeight(int height) {
		this.height = height;
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	public void setWidth(int width) {
		this.width = width;
	}//Method
	
	public String returnFolderName() {
		return getFolderName();
	}
	
}//Class