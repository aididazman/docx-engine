package com.docx.model;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.IBodyElement;

public class DocxVO {

	private IBodyElement bodyElement;
	private Map<String, Object> mapValues;
	private CollectionDO collectionDO;
	private TagDO tagDO;
	private ParentTableDO parentTableDO;
	private boolean isCollectionInTable;


	public IBodyElement getBodyElement() {
		return bodyElement;
	}

	public void setBodyElement(IBodyElement bodyElement) {
		this.bodyElement = bodyElement;
	}

	public Map<String, Object> getMapValues() {
		return mapValues;
	}

	public void setMapValues(Map<String, Object> mapValues) {
		this.mapValues = mapValues;
	}

	public CollectionDO getCollectionDO() {
		return collectionDO;
	}

	public void setCollectionDO(CollectionDO collectionDO) {
		this.collectionDO = collectionDO;
	}

	public TagDO getTagDO() {
		return tagDO;
	}

	public void setTagDO(TagDO tagDO) {
		this.tagDO = tagDO;
	}

	public ParentTableDO getParentTableDO() {
		return parentTableDO;
	}

	public void setParentTableDO(ParentTableDO parentTableDO) {
		this.parentTableDO = parentTableDO;
	}

	public boolean isCollectionInTable() {
		return isCollectionInTable;
	}

	public void setIsCollectionInTable(boolean isCollectionInTable) {
		this.isCollectionInTable = isCollectionInTable;
	}
	
	
}
