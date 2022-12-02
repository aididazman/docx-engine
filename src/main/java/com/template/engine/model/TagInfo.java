package com.template.engine.model;

public class TagInfo {

	private String tagText;
	private int tagStartOffset = 0;
	private boolean hasClosingSlash = false;
	private int tagElementIndex = 0;
	private boolean hasProcessed = false;

	public TagInfo(String tagText, int tagStartOffset, boolean hasClosingSlash) {
		this.tagText = tagText;
		this.tagStartOffset = tagStartOffset;
		this.hasClosingSlash = hasClosingSlash;
	}
	
	public TagInfo(String tagText, int tagStartOffset, boolean hasClosingSlash, int tagElementIndex) {
		this.tagText = tagText;
		this.tagStartOffset = tagStartOffset;
		this.hasClosingSlash = hasClosingSlash;
		this.tagElementIndex = tagElementIndex;
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

	public boolean hasClosingSlash() {
		return hasClosingSlash;
	}

	public void hasClosingSlash(boolean hasClosingSlash) {
		this.hasClosingSlash = hasClosingSlash;
	}

	public int getTagElementIndex() {
		return tagElementIndex;
	}

	public void setTagElementIndex(int tagElementIndex) {
		this.tagElementIndex = tagElementIndex;
	}
	
	public boolean isHasProcessed() {
		return hasProcessed;
	}

	public void setHasProcessed(boolean hasProcessed) {
		this.hasProcessed = hasProcessed;
	}

	@Override
	public String toString() {
		return "TagInfo [tagText=" + tagText + ", tagStartOffset=" + tagStartOffset + ", hasClosingSlash="
				+ hasClosingSlash + ", tagElementIndex=" + tagElementIndex + ", hasProcessed=" + hasProcessed + "]";
	}

}
