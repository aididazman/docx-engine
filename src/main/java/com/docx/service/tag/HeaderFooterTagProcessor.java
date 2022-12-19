package com.docx.service.tag;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import com.docx.service.utils.DocxUtils;

public class HeaderFooterTagProcessor {

	public String processValue(String tagName, Map<String, Object> resolutionAttributesMap) {
		String value = DocxUtils.processValue(tagName, resolutionAttributesMap);
		return value;
	}

	public void fillHeaderFooterTag(XWPFParagraph paragraph, String textToBeReplaced, String textValue) {
		DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
	}

}
