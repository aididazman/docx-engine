package com.template.engine.model;

import org.apache.poi.xwpf.usermodel.XWPFTable;

public class ParentTableDO {

	private XWPFTable table;
	private int rowIndex;
	private int cellIndex;

	public XWPFTable getTable() {
		return table;
	}

	public void setTable(XWPFTable table) {
		this.table = table;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public int getCellIndex() {
		return cellIndex;
	}

	public void setCellIndex(int cellIndex) {
		this.cellIndex = cellIndex;
	}
}
