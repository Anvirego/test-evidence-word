package com.github.anvirego;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
/**
 * @author Ing. Angélica Viridiana Rebolloza Gonzalez.
 * @version 3.0 03/2021.
 * TestEvidence: Creates the WordDocument from scratch.
 */
public final class TestEvidence extends TestEvidenceLogic {
	private static TestEvidence testE;
	private static XWPFDocument wordDocument;
	private int height, width;
	//private static String suiteTestName, platform, testerName, environment;
	//static String folderName;

	public TestEvidence(String platform, String suiteTestName, String testerName, String environment) {
		/*
		TestEvidence.platform = platform;
		TestEvidence.suiteTestName = suiteTestName;
		TestEvidence.testerName = testerName;
		TestEvidence.environment = environment;
		*/
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
	//public XWPFDocument createWordDocument() {
	/*
	private void createWordDocument() {
		System.out.println("::::: createWordDocument :::::");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		Date date = cal.getTime();             
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		//Blank Word Document
		folderName = format.format(date);
		File folder = new File(folderName);
		folder.mkdir(); 
		wordDocument = new XWPFDocument();
		try {
			//Upper Table
			XWPFTable data = wordDocument.createTable(2, 4);//createTable(row, column)
			data.setWidth(10000);
			setTableAlign(data, ParagraphAlignment.CENTER);
			XWPFParagraph p1 = data.getRow(0).getCell(0).getParagraphs().get(0);
			p1.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun r1 = p1.createRun();
			r1.addBreak();
			//r1.addPicture(this.getClass().getResourceAsStream("logo.png"), XWPFDocument.PICTURE_TYPE_PNG, null, Units.toEMU(120), Units.toEMU(27));
			r1.addPicture(TestEvidence.class.getResourceAsStream("logo.png"), XWPFDocument.PICTURE_TYPE_PNG, null, Units.toEMU(120), Units.toEMU(27));
			//Combines cells
			CTHMerge hMerge = CTHMerge.Factory.newInstance();
			hMerge.setVal(STMerge.RESTART);
			data.getRow(0).getCell(1).getCTTc().addNewTcPr().setHMerge(hMerge);
			CTHMerge hMerge1 = CTHMerge.Factory.newInstance();
			hMerge.setVal(STMerge.CONTINUE);
			data.getRow(0).getCell(2).getCTTc().addNewTcPr().setHMerge(hMerge1);
			CTHMerge hMerge2 = CTHMerge.Factory.newInstance();
			hMerge2.setVal(STMerge.CONTINUE);
			data.getRow(0).getCell(3).getCTTc().addNewTcPr().setHMerge(hMerge2);
			cells(data,0,1,"Test Name:",1,true,"blue",10);
			cells(data,0,1,suiteTestName,0,true,"blue",11);
			cells(data,1,0,"Test Evidence",1,true,"blue",9);
			cells(data,1,0,""+format.format(date),0,true,"black",9);
			cells(data,1,1,"Operating System:",1,true,"blue",10);
			cells(data,1,1,platform,0,false,"black",9);
			cells(data,1,2,"Tester Name:",1,true,"blue",10);
			cells(data,1,2,testerName,0,false,"black",9);
			cells(data,1,3,"Environment",1,true,"blue",10);
			cells(data,1,3,environment,0,false,"black",9);
			XWPFParagraph p2 = wordDocument.createParagraph();
			p2.setAlignment(ParagraphAlignment.CENTER);
		} catch (Exception e) {System.out.println("▓▓▓▓▓▓▓▓▓▓ createWordDocument: "+e+" ▓▓▓▓▓▓▓▓▓▓");};
		//return wordDocument;
	}//Method
	*/
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	/*
	private static void setTableAlign(XWPFTable table, ParagraphAlignment align) {
		CTTblPr tblPr = table.getCTTbl().getTblPr();
		CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : (CTJc) tblPr.addNewJc());
		STJc.Enum en = STJc.Enum.forInt(align.getValue());
		jc.setVal(en);
	}//Method
	*/
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	/*
	private static void cells(XWPFTable data, int row, int cell, String text, int salto, boolean bold, String color, int size) {
		try {
			XWPFParagraph p1 = data.getRow(row).getCell(cell).getParagraphs().get(0);
			p1.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun r1 = p1.createRun();
			r1.setFontFamily("Candara");
			r1.setText(text);
			r1.setBold(bold);
			r1.setFontSize(size);
			if(color.equalsIgnoreCase("blue")) {
				r1.setColor("2A0573");
			}
			if(salto == 1) {
				r1.addBreak();
			}
		} catch (Exception e) {System.out.println("▓▓▓▓▓▓▓▓▓▓ cells: "+e+" ▓▓▓▓▓▓▓▓▓▓");};
	}//Method
	*/
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
	/*
	private Boolean signEvidence() {
		XWPFParagraph evidenceP = wordDocument.createParagraph();
		XWPFRun evidenceR = evidenceP.createRun();
		evidenceR.setBold(true);
		evidenceR.setFontSize(5);
		evidenceR.setText("Word created by: Angelica V. Rebolloza G.");
		return true;
	}
	*/
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	/*
	private void deleteDirCMD(String dirName) throws IOException {
		System.out.println("::::: deleteDirCMD :::::");
		System.out.println("::::: Deleting: "+dirName);
		System.out.println("::::: "+System.getProperty("os.name")+" :::::");
		if(System.getProperty("os.name").equalsIgnoreCase("MAC OS X")) {
			System.out.println("::::: Mac OS :::::");
			Process process = Runtime.getRuntime().exec("rm -r "+dirName);
			printResults(process);
		} else {
			System.out.println("::::: Windows OS :::::");
			Process process = Runtime.getRuntime().exec("rmdir "+dirName);
			printResults(process);
		}
	}//Method
	*/
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	/*
	private void printResults(Process process) throws IOException {
	    BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
	    String line = "";
	    while ((line = reader.readLine()) != null) {
	        System.out.println(line);
	    }
	}//Method
	*/
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