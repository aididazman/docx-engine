package com.docx.service.engine;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Main2 {
	
	public static void main(String[] args) throws Exception {

	    XWPFDocument document = new XWPFDocument();

	    XWPFTable tableOne = document.createTable(2,2);
	    XWPFTableRow tablerow = tableOne.getRow(0);
	    tablerow.getCell(0).setText("Test");
	    tablerow.getCell(1).setText("Test");

	    tablerow = tableOne.getRow(1);
	    tablerow.getCell(0).setText("Test");

	    XWPFParagraph paragraph = tablerow.getCell(1).getParagraphArray(0);
	    XWPFTable tableTwo = tablerow.getCell(1).insertNewTbl(paragraph.getCTP().newCursor());

	    tableTwo.getCTTbl().addNewTblPr().addNewTblBorders().addNewLeft().setVal(
	     org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
	    tableTwo.getCTTbl().getTblPr().getTblBorders().addNewRight().setVal(
	     org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
	    tableTwo.getCTTbl().getTblPr().getTblBorders().addNewTop().setVal(
	     org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
	    tableTwo.getCTTbl().getTblPr().getTblBorders().addNewBottom().setVal(
	     org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
	    tableTwo.getCTTbl().getTblPr().getTblBorders().addNewInsideH().setVal(
	     org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
	    tableTwo.getCTTbl().getTblPr().getTblBorders().addNewInsideV().setVal(
	     org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);

	    tablerow = tableTwo.createRow();
	    tablerow.createCell().setText("aaaaaaaaaa");
	    tablerow.createCell().setText("jjjjjjjj"); 
	    tablerow = tableTwo.createRow(); 
	    tablerow.getCell(0).setText("bbbbbbbbbb"); 
	    tablerow.getCell(1).setText("gggggggggg");

	    document.write(new FileOutputStream("CreateWordTableInTable.docx"));
	    document.close();

	 }

}
