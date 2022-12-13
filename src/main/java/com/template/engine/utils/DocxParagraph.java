package com.template.engine.utils;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.template.engine.model.TagInfo;

public class DocxParagraph {

	public void editParagraph(XWPFParagraph paragraph, TagInfo tag, String value) {
		
		String textToBeReplaced = DocxUtils.addTagBracket(tag.getTagText());
		if(!DocxUtils.isNullEmpty(paragraph.getText())) {
			if(paragraph.getText().contains(textToBeReplaced)) {
				DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, value);
			}
		}
	}
	
	public XWPFParagraph createNewParagraph(XWPFParagraph parentParagraph, XWPFParagraph endParagraph, TagInfo tag, String value) {
		
		XWPFParagraph newParagraph = null;
		
		String textToBeReplaced = DocxUtils.addTagBracket(tag.getTagText());
		if(!DocxUtils.isNullEmpty(parentParagraph.getText())) {
			if(parentParagraph.getText().contains(textToBeReplaced)) {
				XWPFDocument document = parentParagraph.getDocument();
				newParagraph = document.insertNewParagraph(endParagraph.getCTP().newCursor());
				XWPFRun run = newParagraph.createRun();
				run.setText(parentParagraph.getText());
				DocxUtils.replaceTextSegment(newParagraph, textToBeReplaced, value);
			}
		}
		
		return newParagraph;
	}
}
