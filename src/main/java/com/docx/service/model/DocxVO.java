package com.docx.service.model;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.IBodyElement;

public class DocxVO {

	private IBodyElement bodyElement;
	private Map<String, Object> resolutionAttributesMap;
	private CollectionDO collectionDO;
	private TagInfo tag;
	private ParentTableDO parentTableDO;

	public IBodyElement getBodyElement() {
		return bodyElement;
	}

	public void setBodyElement(IBodyElement bodyElement) {
		this.bodyElement = bodyElement;
	}

	public Map<String, Object> getResolutionAttributesMap() {
		return resolutionAttributesMap;
	}

	public void setResolutionAttributesMap(Map<String, Object> resolutionAttributesMap) {
		this.resolutionAttributesMap = resolutionAttributesMap;
	}

	public CollectionDO getCollectionDO() {
		return collectionDO;
	}

	public void setCollectionDO(CollectionDO collectionDO) {
		this.collectionDO = collectionDO;
	}

	public TagInfo getTag() {
		return tag;
	}

	public void setTag(TagInfo tag) {
		this.tag = tag;
	}

	public ParentTableDO getParentTableDO() {
		return parentTableDO;
	}

	public void setParentTableDO(ParentTableDO parentTableDO) {
		this.parentTableDO = parentTableDO;
	}
	
	
}
