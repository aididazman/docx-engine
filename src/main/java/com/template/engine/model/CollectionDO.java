package com.template.engine.model;

import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.IBodyElement;

public class CollectionDO {

	private String mapKey;
	private String tagName;
	private Map<String, Object> resolutionAttributesMap;
	private List<Object> collectionValues;
	private int startCollectionIndex;
	private String startCollectionName;
	private Integer endCollectionIndex;
	private String endCollectionName;
	private IBodyElement endCollectionElement;
	private boolean isLastCollectionValue;
	private boolean isNestedCollection;
	private boolean isElementInTable;
	private ParentTableDO parentTableDO;

	public String getMapKey() {
		return mapKey;
	}

	public void setMapKey(String mapKey) {
		this.mapKey = mapKey;
	}

	public String getTagName() {
		return tagName;
	}

	public void setTagName(String tagName) {
		this.tagName = tagName;
	}

	public Map<String, Object> getResolutionAttributesMap() {
		return resolutionAttributesMap;
	}

	public void setResolutionAttributesMap(Map<String, Object> resolutionAttributesMap) {
		this.resolutionAttributesMap = resolutionAttributesMap;
	}

	public List<Object> getCollectionValues() {
		return collectionValues;
	}

	public void setCollectionValues(List<Object> collectionValues) {
		this.collectionValues = collectionValues;
	}

	public int getStartCollectionIndex() {
		return startCollectionIndex;
	}

	public void setStartCollectionIndex(int startCollectionIndex) {
		this.startCollectionIndex = startCollectionIndex;
	}

	public Integer getEndCollectionIndex() {
		return endCollectionIndex;
	}

	public void setEndCollectionIndex(Integer endCollectionIndex) {
		this.endCollectionIndex = endCollectionIndex;
	}

	public IBodyElement getEndCollectionElement() {
		return endCollectionElement;
	}

	public void setEndCollectionElement(IBodyElement endCollectionElement) {
		this.endCollectionElement = endCollectionElement;
	}

	public boolean isLastCollectionValue() {
		return isLastCollectionValue;
	}

	public void setLastCollectionValue(boolean isLastCollectionValue) {
		this.isLastCollectionValue = isLastCollectionValue;
	}

	public boolean isNestedCollection() {
		return isNestedCollection;
	}

	public void setNestedCollection(boolean isNestedCollection) {
		this.isNestedCollection = isNestedCollection;
	}

	public boolean isElementInTable() {
		return isElementInTable;
	}

	public void setElementInTable(boolean isElementInTable) {
		this.isElementInTable = isElementInTable;
	}

	public ParentTableDO getParentTableDO() {
		return parentTableDO;
	}

	public void setParentTableDO(ParentTableDO parentTableDO) {
		this.parentTableDO = parentTableDO;
	}

	public String getStartCollectionName() {
		return startCollectionName;
	}

	public void setStartCollectionName(String startCollectionName) {
		this.startCollectionName = startCollectionName;
	}

	public String getEndCollectionName() {
		return endCollectionName;
	}

	public void setEndCollectionName(String endCollectionName) {
		this.endCollectionName = endCollectionName;
	}

}
