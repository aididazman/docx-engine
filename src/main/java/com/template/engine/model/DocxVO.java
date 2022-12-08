package com.template.engine.model;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class DocxVO {

	private XWPFParagraph paragraphBodyElement;
	private XWPFTableCell tableCell;
	private boolean isTagInTableCell;
	private int elementIndex;

	public XWPFParagraph getParagraphBodyElement() {
		return paragraphBodyElement;
	}

	public void setParagraphBodyElement(XWPFParagraph paragraphBodyElement) {
		this.paragraphBodyElement = paragraphBodyElement;
	}

	public XWPFTableCell getTableCell() {
		return tableCell;
	}

	public void setTableCell(XWPFTableCell tableCell) {
		this.tableCell = tableCell;
	}

	public boolean isTagInTableCell() {
		return isTagInTableCell;
	}

	public void setTagInTableCell(boolean isTagInTableCell) {
		this.isTagInTableCell = isTagInTableCell;
	}

	public int getElementIndex() {
		return elementIndex;
	}

	public void setElementIndex(int elementIndex) {
		this.elementIndex = elementIndex;
	}

}
