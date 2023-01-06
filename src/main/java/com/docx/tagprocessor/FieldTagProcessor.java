package com.docx.tagprocessor;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import com.docx.utils.DocxUtils;

public class FieldTagProcessor {
	
	public void fillTag(XWPFParagraph paragraph, String textToBeReplaced, String textValue) {
		
		if (!DocxUtils.isNullEmpty(textValue))
			DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
	}

	public String processValue(String tagName, Map<String, Object> resolutionAttributesMap) {
		String value = DocxUtils.processValue(tagName, resolutionAttributesMap);
		return value;
	}

}
