package com.template.engine.model;

import org.apache.poi.xwpf.usermodel.IBodyElement;

public class TagInfo {

	private String tagText;
	private int tagStartOffset = 0;
	private IBodyElement tagElement;
	
	public TagInfo(String tagText, int tagStartOffset, IBodyElement tagElement) {
		super();
		this.tagText = tagText;
		this.tagStartOffset = tagStartOffset;
		this.tagElement = tagElement;
	}

	public String getTagText() {
		return tagText;
	}

	public void setTagText(String tagText) {
		this.tagText = tagText;
	}

	public int getTagStartOffset() {
		return tagStartOffset;
	}

	public void setTagStartOffset(int tagStartOffset) {
		this.tagStartOffset = tagStartOffset;
	}

	public IBodyElement getTagElement() {
		return tagElement;
	}

	public void setTagElement(IBodyElement tagElement) {
		this.tagElement = tagElement;
	}
	
	@Override
	public String toString() {
		return "TagInfo [tagText=" + tagText + ", tagElement=" + tagElement + "]";
	}

}
