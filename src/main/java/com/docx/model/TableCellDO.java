package com.docx.model;

import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class TableCellDO {

	private XWPFTableCell cellTable;
	private IBodyElement newCellBodyElement;
	private TagDO tagDO;
	private CollectionDO collectionDO;
	private int rowIndex;
	private List<IBodyElement> parentCellBodyElements;
	private IBodyElement parentCellBodyElement;
	private Map<String, Object> resolutionAttributesMap;

	public XWPFTableCell getCellTable() {
		return cellTable;
	}

	public void setCellTable(XWPFTableCell cellTable) {
		this.cellTable = cellTable;
	}

	public IBodyElement getNewCellBodyElement() {
		return newCellBodyElement;
	}

	public void setNewCellBodyElement(IBodyElement newCellBodyElement) {
		this.newCellBodyElement = newCellBodyElement;
	}

	public TagDO getTagDO() {
		return tagDO;
	}

	public void setTagDO(TagDO tagDO) {
		this.tagDO = tagDO;
	}

	public CollectionDO getCollectionDO() {
		return collectionDO;
	}

	public void setCollectionDO(CollectionDO collectionDO) {
		this.collectionDO = collectionDO;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public List<IBodyElement> getParentCellBodyElements() {
		return parentCellBodyElements;
	}

	public void setParentCellBodyElements(List<IBodyElement> parentCellBodyElements) {
		this.parentCellBodyElements = parentCellBodyElements;
	}

	public IBodyElement getParentCellBodyElement() {
		return parentCellBodyElement;
	}

	public void setParentCellBodyElement(IBodyElement parentCellBodyElement) {
		this.parentCellBodyElement = parentCellBodyElement;
	}

	public Map<String, Object> getResolutionAttributesMap() {
		return resolutionAttributesMap;
	}

	public void setResolutionAttributesMap(Map<String, Object> resolutionAttributesMap) {
		this.resolutionAttributesMap = resolutionAttributesMap;
	}
}
