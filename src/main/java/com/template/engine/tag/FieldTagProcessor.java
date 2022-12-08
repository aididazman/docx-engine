package com.template.engine.tag;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.template.engine.utils.DocxUtils;

public class FieldTagProcessor {

	public void fillFieldTag(IBodyElement bodyElem, String textToBeReplaced, String textValue) {

		if (bodyElem instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
		} else if (bodyElem instanceof XWPFTable) {
			XWPFTable table = (XWPFTable) bodyElem;
			for (XWPFTableRow row : table.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph paragraph : cell.getParagraphs()) {
						DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
					}
				}
			}
		}
	}

	public String processValue(String tagName, Map<String, Object> resolutionAttributesMap) {
		String value = DocxUtils.processValue(tagName, resolutionAttributesMap);
		return value;
	}

}
