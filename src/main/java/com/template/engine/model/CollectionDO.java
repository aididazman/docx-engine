package com.template.engine.model;

import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.IBodyElement;

public class CollectionDO {

	private String objectFirstParameter;
	private String objectSecondParameter;
	private String tagName;
	private Map<String, Object> resolutionAttributesMap;
	private TagInfo tag;
	private List<Object> collectionValues;
	private int startCollectionIndex;
	private Integer endCollectionIndex;
	private IBodyElement endCollectionElement;
	private boolean hasNestedCollection;
	private boolean isLastCollectionValue;

	public String getObjectFirstParameter() {
		return objectFirstParameter;
	}

	public void setObjectFirstParameter(String objectFirstParameter) {
		this.objectFirstParameter = objectFirstParameter;
	}

	public String getObjectSecondParameter() {
		return objectSecondParameter;
	}

	public void setObjectSecondParameter(String objectSecondParameter) {
		this.objectSecondParameter = objectSecondParameter;
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

	public TagInfo getTag() {
		return tag;
	}

	public void setTag(TagInfo tag) {
		this.tag = tag;
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

	public boolean isHasNestedCollection() {
		return hasNestedCollection;
	}

	public void setHasNestedCollection(boolean hasNestedCollection) {
		this.hasNestedCollection = hasNestedCollection;
	}

	public boolean isLastCollectionValue() {
		return isLastCollectionValue;
	}

	public void setLastCollectionValue(boolean isLastCollectionValue) {
		this.isLastCollectionValue = isLastCollectionValue;
	}

}
