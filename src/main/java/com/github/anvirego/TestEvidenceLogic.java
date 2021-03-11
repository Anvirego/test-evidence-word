package com.github.anvirego;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.util.Units;
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

import com.github.anvirego.interfaces.TestEvidenceInterface;

/**
 * @author Ing. Angelica Viridiana Rebolloza Gonzalez.
 * @version 2.0 03/2021.
 * TestEvidenceLogic: Library main logic. 
 */
public class TestEvidenceLogic implements TestEvidenceInterface{
	private static String folderName;
	private static XWPFDocument wordDocument;
	
	protected XWPFDocument createWordDocument(String platform, String suiteTestName, String testerName, String environment) {
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
			r1.addPicture(new FileInputStream(new File(System.getProperty("user.dir").concat("/src/test/resources/Logo/logo.png"))), XWPFDocument.PICTURE_TYPE_PNG, null, Units.toEMU(120), Units.toEMU(27));
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
		return wordDocument;
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected static void setTableAlign(XWPFTable table, ParagraphAlignment align) {
		CTTblPr tblPr = table.getCTTbl().getTblPr();
		CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : (CTJc) tblPr.addNewJc());
		STJc.Enum en = STJc.Enum.forInt(align.getValue());
		jc.setVal(en);
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected static void cells(XWPFTable data, int row, int cell, String text, int salto, boolean bold, String color, int size) {
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
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
	protected Boolean signEvidence() {
		XWPFParagraph evidenceP = wordDocument.createParagraph();
		XWPFRun evidenceR = evidenceP.createRun();
		evidenceR.setBold(true);
		evidenceR.setFontSize(5);
		evidenceR.setText("Word created by: Angelica V. Rebolloza G.");
		return true;
	}
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	protected void deleteDirCMD(String dirName) throws IOException {
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
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	protected void printResults(Process process) throws IOException {
	    BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
	    String line = "";
	    while ((line = reader.readLine()) != null) {
	        System.out.println(line);
	    }
	}//Method
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	protected String getFolderName() {
		return folderName;
	}
//▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄	
	public void createEvidence(String path, String txtDescription, int step) {
		// TODO Auto-generated method stub	
	}

	public Boolean closeWord(String path, FileOutputStream outPicture) {
		// TODO Auto-generated method stub
		return null;
	}

	public void setHeight(int height) {
		// TODO Auto-generated method stub	
	}
	
	public void setWidth(int width) {
		// TODO Auto-generated method stub	
	}

	public String returnFolderName() {
		// TODO Auto-generated method stub
		return null;
	}

}//Class