package com.docx.tagprocessor;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections.IteratorUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import com.docx.model.CollectionDO;
import com.docx.model.ParentTableDO;
import com.docx.model.TableCellDO;
import com.docx.model.TagDO;
import com.docx.utils.DocxConstants;
import com.docx.utils.DocxUtils;

public class CollectionTagProcessor {
	
	private static final Pattern OBJECT_FIELD_PATTERN_1 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+");
	private static final Pattern OBJECT_FIELD_PATTERN_2 = Pattern.compile("[a-zA-Z]+:[a-zA-Z]+");
	
	private static final Pattern COLLECTION_START_PATTERN_1 = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_2 = Pattern.compile("\\$\\{collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	
	private static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	
	private static final Pattern COLLECTION_END_PATTERN_1 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_2 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");

	public CollectionDO getCollection(CollectionDO collectionDO) throws Exception {
		collectionDO = getCollectionValues(collectionDO);
		return collectionDO;
	}

	public void newProcess(CollectionDO collectionDO) throws Exception {
		
		if (collectionDO.getNextElement() instanceof XWPFParagraph) {
			processCollectionInParagraph(collectionDO);
		} else if (collectionDO.getNextElement() instanceof XWPFTable) {
			processCollectionInTable(collectionDO);
		}
	}

	private List<Object> getCollectionValue(String objectName, Object mapValue) {

		Collection<?> items = null;
		try {
			items = (Collection<?>) PropertyUtils.getSimpleProperty(mapValue, objectName);
		} catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
			throw new RuntimeException("Cannot get object " + objectName + " value from " + mapValue);
		}

		List<Object> listOfItems = new ArrayList<>(items);

		return listOfItems;
	}

	private CollectionDO getCollectionValues(CollectionDO collectionDO) throws Exception {

		List<Object> collectionValues = null;
		String collectionObjectField = null;
		
		// check type pattern
		if (OBJECT_FIELD_PATTERN_1.matcher(collectionDO.getTagName()).matches()
				|| OBJECT_FIELD_PATTERN_2.matcher(collectionDO.getTagName()).matches()) {
			
			collectionObjectField = DocxUtils.getObjectField(collectionDO.getMapKey());
			if (collectionDO.getNewMap() != null) {
				collectionValues = getCollectionValue(collectionObjectField,
						collectionDO.getNewMap().get(collectionDO.getMapKey()));
			} else {
				collectionValues = getCollectionValue(collectionObjectField,
						collectionDO.getResolutionAttributesMap().get(collectionDO.getMapKey()));
			}
			collectionDO.setCollectionValues(collectionValues);
		}

		if (collectionValues == null)
			throw new Exception ("Empty collection values");
		
		return collectionDO;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCollectionInParagraph(CollectionDO collectionDO) throws Exception {

		List<TagDO> tagDOList = new ArrayList<>();
		String nonParentNestedCollectionKey = null;
		XWPFParagraph paragraph = (XWPFParagraph) collectionDO.getNextElement();
		Map<String, Object> resolutionAttributesMap = collectionDO.getResolutionAttributesMap();
		
		// throw new error here if no end collection index found
		tagDOList = getTagsFromCollection(collectionDO.getNextElement(), tagDOList,
				collectionDO.getStartCollectionIndex(), collectionDO.getEndCollectionIndex());

		for (Object collectionValue : collectionDO.getCollectionValues()) {

			for (TagDO tagDO : tagDOList) {
				String tagText = DocxUtils.addTagBracket(tagDO.getTagText());

				if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
					//example value listOfUser:user, get field -> user
					String collectionName = DocxUtils.getObjectField(collectionDO.getTagName());
					// example value from user.name -> user
					String objectName = DocxUtils.getObjectName(tagDO.getTagText());
					if (collectionName.equals(objectName)) {
						// example value from user.name -> name
						String objectField = DocxUtils.getObjectField(tagDO.getTagText());
						// get the value of the field based on the tag object field
						String value = DocxUtils.getFieldValue(objectField, collectionValue);
						// replace tag text with
						insertNewParagraph(paragraph, value, collectionDO, tagDO);
					}
				}

				else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
						|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
					String tagName = DocxUtils.getTagName(tagDO.getTagText(), DocxConstants.TAG_PREFIX_COLLECTION_START);

					if (OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
						// returns example value user.phones:phone -> user.phones
						String nestedCollectionKey = DocxUtils.getMapKey(tagName);
						// returns user from user.phones
						String nestedCollectionName = DocxUtils.getObjectName(nestedCollectionKey);
						//example value listOfUser:user, get field -> user
						String parentCollectionName = DocxUtils.getObjectField(collectionDO.getTagName());
						// belongs to parent collection values
						if (parentCollectionName.equals(nestedCollectionName)) {
							// new map
							Map<String, Object> newValues = new HashMap<String, Object>();
							newValues.put(nestedCollectionKey, collectionValue);
							
							CollectionDO nestedCollectionDO = new CollectionDO();
							boolean isLastCollectionValue = collectionValue == collectionDO.getCollectionValues()
									.get(collectionDO.getCollectionValues().size() - 1);
							if (isLastCollectionValue) {
								nestedCollectionDO.setLastCollectionValue(true);
							}
							nestedCollectionDO.setTagName(tagName);
							nestedCollectionDO.setMapKey(nestedCollectionKey);
							nestedCollectionDO.setNewMap(newValues);
							nestedCollectionDO.setResolutionAttributesMap(collectionDO.getResolutionAttributesMap());
							nestedCollectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(tagDO.getTagElement()));
							nestedCollectionDO.setNestedCollection(true);
							nestedCollectionDO = getEndCollection(paragraph, DocxUtils.getElementIndex(tagDO.getTagElement()),
									nestedCollectionDO, collectionDO.getEndCollectionElement());	
							nestedCollectionDO = getCollection(nestedCollectionDO);
							if (collectionDO.isElementInTable()) {
								nestedCollectionDO.setElementInTable(true);
								nestedCollectionDO.setParentTableDO(collectionDO.getParentTableDO());
							}
							
							IBodyElement nestedCollectionElement = tagDO.getTagElement();
							
							if (nestedCollectionElement instanceof XWPFParagraph) {
								XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
								nestedCollectionDO.setNextElement(DocxUtils.getNextElement(nestedParagraph));
								newProcess(nestedCollectionDO);
							} 
						}
					}

					else if (OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
						nonParentNestedCollectionKey = DocxUtils.getMapKey(tagName); // employees:name -> employees
						boolean isProcessed = isCollectionProcessed(tagDOList, tagDO);
						
						if(!isProcessed) {
							
							if (resolutionAttributesMap.containsKey(nonParentNestedCollectionKey)) {
								
								List<Object> newCollectionValues = null;
								if (resolutionAttributesMap.get(nonParentNestedCollectionKey) instanceof ArrayList) {
									ListIterator<Object> iterator = ((ArrayList) resolutionAttributesMap
											.get(nonParentNestedCollectionKey)).listIterator();
									newCollectionValues = IteratorUtils.toList(iterator);
								}
								
								Map<String, Object> newValues = new HashMap<String, Object>();
								newValues.put(nonParentNestedCollectionKey, newCollectionValues);
								
								CollectionDO nestedCollectionDO = new CollectionDO();		
								boolean isLastCollectionValue = collectionValue == collectionDO
										.getCollectionValues().get(collectionDO.getCollectionValues().size() - 1);
								if (isLastCollectionValue) {
									nestedCollectionDO.setLastCollectionValue(true);
								}						
								nestedCollectionDO.setCollectionValues(newCollectionValues);
								nestedCollectionDO.setTagName(tagName);
								nestedCollectionDO.setMapKey(nonParentNestedCollectionKey);
								nestedCollectionDO.setNewMap(newValues);
								nestedCollectionDO.setResolutionAttributesMap(resolutionAttributesMap);
								nestedCollectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(tagDO.getTagElement()));
								nestedCollectionDO.setNestedCollection(true);								
								nestedCollectionDO = getEndCollection(paragraph, DocxUtils.getElementIndex(tagDO.getTagElement()),
										nestedCollectionDO, collectionDO.getEndCollectionElement());
								if (collectionDO.isElementInTable()) {
									nestedCollectionDO.setElementInTable(true);
									nestedCollectionDO.setParentTableDO(collectionDO.getParentTableDO());
								}

								IBodyElement nestedCollectionElement = tagDO.getTagElement();
								
								if (nestedCollectionElement instanceof XWPFParagraph) {
									XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
									nestedCollectionDO.setNextElement(DocxUtils.getNextElement(nestedParagraph));
									newProcess(nestedCollectionDO);
								}	
							}
						}
					}
				}
			}
		}	
	}
	
	private boolean isCollectionProcessed(List<TagDO> tagDOList, TagDO tagDOText)
			throws Exception {
		
		List<TagDO> tagDOSubList = new ArrayList<>();
		
		//get sublist of tag
		for(TagDO tagDO : tagDOList) {
			if(!tagDOText.getTagText().equals(tagDO.getTagText())) {
				tagDOSubList.add(tagDO);
			} else
				break;
		}
		
		Collections.reverse(tagDOSubList);
		for (TagDO subListTagDO : tagDOSubList) {
			String tagText = DocxUtils.addTagBracket(subListTagDO.getTagText());
			if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
					|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
				return true;
			}
		}
		
		return false;
	}

	private void processNestedCollectionTableInParagraph(XWPFTable tableElementAfterStartCollection,
			CollectionDO collectionDO) throws Exception {
		
		XWPFTable parentTable = (XWPFTable) tableElementAfterStartCollection;
		
		IBodyElement endCollectionElement = collectionDO.getEndCollectionElement();
		XWPFTable newTable = insertNewTable(parentTable, endCollectionElement);
		newTable.setWidth(parentTable.getWidth());
		
		if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {
			
			int newRowSize = collectionDO.getCollectionValues().size();
			
			for (int cell = 1; cell < parentTable.getRows().get(0).getTableCells().size(); cell++) {
				newTable.getRow(0).createCell();
			}
			
			for (int row = 0; row < newRowSize; row++) {
				newTable.createRow();
			}

			for (int row = 0; row < newTable.getRows().size(); row++) {
				XWPFTableRow newTableRow = newTable.getRows().get(row);
				
				for (int cell = 0; cell < newTableRow.getTableCells().size(); cell++) {
					XWPFTableCell newTableCell = newTableRow.getTableCells().get(cell);
					IBodyElement newCellBodyElem = newTableCell.getBodyElements().get(0);
					// to get the tags in the first row, get the cell's body element
					List<IBodyElement> parentCellBodyElements = parentTable.getRow(1).getCell(cell).getBodyElements();
					
					if (row == 0) {
						if (newCellBodyElem instanceof XWPFParagraph) {
							XWPFParagraph paragraph = (XWPFParagraph) newCellBodyElem;
							if (DocxUtils.isNullEmpty(paragraph.getText())) {
								XWPFRun run = paragraph.createRun();
								String text = run.getText(0);
								
								if (DocxUtils.isNullEmpty(text))
									text = DocxConstants.EMPTY_STRING;
								
								String headerText = parentTable.getRow(0).getCell(cell).getTextRecursively(); // set in value to be replaced
								text = text.replace(text, headerText);
								run.setText(text, 0);
							}		
						}	
					}
					
					else {
						// loop each element to get the tags in each cell
						for (IBodyElement parentCellBodyElem : parentCellBodyElements) {
							List<TagDO> tagDOList = new ArrayList<>();
							
							if (parentCellBodyElem instanceof XWPFParagraph) {
								XWPFParagraph parentCellParagraph = (XWPFParagraph) parentCellBodyElem;
								String paragraphText = parentCellParagraph.getText();

								if (!DocxUtils.isNullEmpty(paragraphText)) {
									tagDOList = DocxUtils.getTagsByElement(paragraphText, 0, tagDOList, parentCellBodyElem);
									
									for (TagDO tagDO : tagDOList) {
										
										TableCellDO tableCellDO = new TableCellDO();
										tableCellDO.setCellTable(newTableCell);
										tableCellDO.setNewCellBodyElement(newCellBodyElem);
										tableCellDO.setTagDO(tagDO);
										tableCellDO.setCollectionDO(collectionDO);
										tableCellDO.setRowIndex(row+1);
										tableCellDO.setParentCellBodyElements(parentCellBodyElements);
										tableCellDO.setParentCellBodyElement(parentCellBodyElem);
										tableCellDO.setResolutionAttributesMap(collectionDO.getResolutionAttributesMap());
										
										processCell(tableCellDO);
									}
								}
							} 
						}
					}					
				}			
			}			
		}
		
		if (collectionDO.isLastCollectionValue()) {
			parentTable.getBody().getXWPFDocument().removeBodyElement(DocxUtils.getElementIndex(parentTable));
		}
	}
	
	private void processCollectionInTable(CollectionDO collectionDO) throws Exception {

		XWPFTable table = (XWPFTable) collectionDO.getNextElement();
		
		// if nested collection in paragraph was a table
		if (collectionDO.isNestedCollection()) {
			processNestedCollectionTableInParagraph(table, collectionDO);
		}
		
		else {
			// collection in table
			for (int value = 0; value < collectionDO.getCollectionValues().size(); value++) {
				table.createRow();
			}
			
			for (int rowIndex = 2; rowIndex < table.getRows().size(); rowIndex++) {

				XWPFTableRow rowTable = table.getRows().get(rowIndex);

				for (int cellIndex = 0; cellIndex < rowTable.getTableCells().size(); cellIndex++) {
					XWPFTableCell cellTable = rowTable.getTableCells().get(cellIndex);
					// to get the tags in the first row, get the cell's body element
					List<IBodyElement> parentCellBodyElements = table.getRow(1).getCell(cellIndex).getBodyElements();
					
					// loop each element to get the tags in each cell
					for (IBodyElement parentCellBodyElem : parentCellBodyElements) {
						List<TagDO> tagDOList = new ArrayList<>();
						
						if (parentCellBodyElem instanceof XWPFParagraph) {
							XWPFParagraph parentCellParagraph = (XWPFParagraph) parentCellBodyElem;
							String paragraphText = parentCellParagraph.getText();

							if (!DocxUtils.isNullEmpty(paragraphText)) {
								tagDOList = DocxUtils.getTagsByElement(paragraphText, 0, tagDOList, parentCellBodyElem);
								
								IBodyElement newCellBodyElem = cellTable.getBodyElements().get(0);
								for (TagDO tagDO : tagDOList) {
									
									TableCellDO tableCellDO = new TableCellDO();
									tableCellDO.setCellTable(cellTable);
									tableCellDO.setNewCellBodyElement(newCellBodyElem);
									tableCellDO.setTagDO(tagDO);
									tableCellDO.setCollectionDO(collectionDO);
									tableCellDO.setRowIndex(rowIndex);
									tableCellDO.setParentCellBodyElements(parentCellBodyElements);
									tableCellDO.setParentCellBodyElement(parentCellBodyElem);
									tableCellDO.setResolutionAttributesMap(collectionDO.getResolutionAttributesMap());
									
									processCell(tableCellDO);				
								}
							}
						} 
					}
				}
			}

			table.removeRow(1);			
		}
	}

	private void processCell(TableCellDO tableCellDO) throws Exception {
		
		TagDO tagDO = tableCellDO.getTagDO();
		CollectionDO collectionDO = tableCellDO.getCollectionDO();
		XWPFTableCell cellTable = tableCellDO.getCellTable();
		String tagText = DocxUtils.addTagBracket(tagDO.getTagText());
		
		if (tableCellDO.getNewCellBodyElement() instanceof XWPFParagraph) {

			if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {

				XWPFParagraph paragraph = (XWPFParagraph) tableCellDO.getNewCellBodyElement();
				String text = paragraph.getText();

				if (DocxUtils.isNullEmpty(text)) {
					String collectionName = DocxUtils.getObjectName(tagDO.getTagText());
					String parentCollectionName = DocxUtils.getObjectField(collectionDO.getTagName());

					if (parentCollectionName.equals(collectionName)) {

						XWPFParagraph parentParagraph = (XWPFParagraph) tableCellDO.getParentCellBodyElement();
						String parentParagraphText = parentParagraph.getText();
						text = text.replace(text, parentParagraphText);
						XWPFRun run = paragraph.createRun();
						run.setText(parentParagraphText, 0);

						Object collectionValue = collectionDO.getCollectionValues().get(tableCellDO.getRowIndex() - 2);

						String objectField = DocxUtils.getObjectField(tagDO.getTagText());
						String value = DocxUtils.getFieldValue(objectField, collectionValue);

						DocxUtils.replaceTextSegment(paragraph, tagText, value);
					}
				}

				else {
					String collectionName = DocxUtils.getObjectName(tagDO.getTagText());
					String parentCollectionName = DocxUtils.getObjectField(collectionDO.getTagName());

					if (parentCollectionName.equals(collectionName)) {

						Object collectionValue = collectionDO.getCollectionValues().get(tableCellDO.getRowIndex() - 2);
						XWPFParagraph newParagraph = cellTable.addParagraph();
						XWPFRun newRun = newParagraph.createRun();
						String newText = newRun.getText(0);

						if (DocxUtils.isNullEmpty(newText))
							newText = DocxConstants.EMPTY_STRING;

						String objectField = DocxUtils.getObjectField(tagDO.getTagText());
						String value = DocxUtils.getFieldValue(objectField, collectionValue);

						XWPFParagraph parentParagraph = (XWPFParagraph) tableCellDO.getParentCellBodyElement();
						String parentParagraphText = parentParagraph.getText();
						newText = newText.replace(newText, parentParagraphText);
						newRun.setText(newText, 0);

						DocxUtils.replaceTextSegment(paragraph, tagText, value);
					}
				}
			}

			else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
					|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {

				String tagName = DocxUtils.getTagName(tagDO.getTagText(),
						DocxConstants.TAG_PREFIX_COLLECTION_START);
				String nestedCollectionKey = DocxUtils.getMapKey(tagName);

				CollectionDO nestedCollectionDO = new CollectionDO();
				nestedCollectionDO.setTagName(tagName);
				nestedCollectionDO.setMapKey(nestedCollectionKey);
				nestedCollectionDO.setResolutionAttributesMap(collectionDO.getResolutionAttributesMap());
				nestedCollectionDO.setStartCollectionIndex(
						DocxUtils.getElementIndex(tableCellDO.getNewCellBodyElement()));
				nestedCollectionDO.setNestedCollection(true);
				nestedCollectionDO = getEndCollection(tableCellDO.getParentCellBodyElement(),
						DocxUtils.getElementIndex(tagDO.getTagElement()), nestedCollectionDO,
						collectionDO.getEndCollectionElement());

				List<TagDO> tagsInBetweenCollection = new ArrayList<>();
				tagsInBetweenCollection = getTagsFromCollection(
						tableCellDO.getParentCellBodyElement(), tagsInBetweenCollection,
						DocxUtils.getElementIndex(tableCellDO.getNewCellBodyElement()),
						nestedCollectionDO.getEndCollectionIndex());

				Object collectionValue = collectionDO.getCollectionValues().get(tableCellDO.getRowIndex() - 2);

				IBodyElement parentCellNextElement = DocxUtils.getNextElement(tableCellDO.getParentCellBodyElement());

				if (parentCellNextElement instanceof XWPFParagraph) {
					processCollectionParagraphInCell(tagName, collectionDO, collectionValue, cellTable,
							tagsInBetweenCollection, nestedCollectionDO);
				}

				else if (parentCellNextElement instanceof XWPFTable) {
					processCollectionTableInCell(parentCellNextElement, tagName, collectionDO, collectionValue,
							cellTable, tagsInBetweenCollection, nestedCollectionDO);
				}
			}
		}
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCollectionParagraphInCell(String tagName, CollectionDO collectionDO,
			Object collectionValue, XWPFTableCell cellTable, List<TagDO> tagsInBetweenCollection,
			CollectionDO nestedCollectionDO) {

		if (OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
			// returns example value user.phones:phone -> user.phones
			String nestedCollectionName = DocxUtils.getMapKey(tagName);
			// returns user from user.phones
			String nestedCollectionObject = DocxUtils.getObjectName(nestedCollectionName);
			//example value listOfUser:user, get field -> user
			String nestedCollectionParentObject = DocxUtils.getObjectField(collectionDO.getTagName());
			
			if (nestedCollectionObject.equals(nestedCollectionParentObject)) {
				String nestedCollectionObjectField = DocxUtils.getObjectField(nestedCollectionName); // phones
				List<Object> nestedCollectionValues = getCollectionValue(nestedCollectionObjectField, collectionValue);
				
				for (Object nestedCollectionValue : nestedCollectionValues) {

					for (TagDO nestedTag : tagsInBetweenCollection) {
						
						String nestedTagText = DocxUtils.addTagBracket(nestedTag.getTagText());
						
						if (COLLECTION_OBJECT_PATTERN.matcher(nestedTagText).matches()) {
							insertParagraphInTableCell(nestedTag, nestedCollectionValue, cellTable);
						}
					}
				}
			}
		}
		
		else if (OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
			String nonParentNestedCollectionName = DocxUtils.getMapKey(tagName); 
			
			if(nestedCollectionDO.getResolutionAttributesMap().containsKey(nonParentNestedCollectionName)) {
				
				if (nestedCollectionDO.getResolutionAttributesMap().get(nonParentNestedCollectionName) instanceof ArrayList) {
					ListIterator<Object> iterator = ((ArrayList)collectionDO.getResolutionAttributesMap()
							.get(nonParentNestedCollectionName)).listIterator();
					List<Object> newCollectionValues = IteratorUtils.toList(iterator);
					
					for (Object nestedCollectionValue : newCollectionValues) {

						for (TagDO nestedTag : tagsInBetweenCollection) {
							
							String nestedTagText = DocxUtils.addTagBracket(nestedTag.getTagText());
							
							if (COLLECTION_OBJECT_PATTERN.matcher(nestedTagText).matches()) {
								insertParagraphInTableCell(nestedTag, nestedCollectionValue, cellTable);
							}
						}
					}															
				}
			}
		}
		
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCollectionTableInCell(IBodyElement firstRowCellNextElement, String tagName, CollectionDO collectionDO,
			Object collectionValue, XWPFTableCell cellTable, List<TagDO> tagsInBetweenCollection,
			CollectionDO nestedCollectionDO) {
		
		XWPFTable nestedFirstRowTable = (XWPFTable) firstRowCellNextElement;
		
		if (OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
			// returns example value user.phones:phone -> user.phones
			String nestedCollectionName = DocxUtils.getMapKey(tagName);
			// returns user from user.phones
			String nestedCollectionObject = DocxUtils.getObjectName(nestedCollectionName);
			//example value listOfUser:user, get field -> user
			String nestedCollectionParentObject = DocxUtils.getObjectField(collectionDO.getTagName());
			
			if (nestedCollectionObject.equals(nestedCollectionParentObject)) {
				String nestedCollectionObjectField = DocxUtils.getObjectField(nestedCollectionName); // phones
				List<Object> nestedCollectionValues = getCollectionValue(nestedCollectionObjectField, collectionValue);
				
				int newRowSize = nestedCollectionValues.size();
				
				XWPFParagraph lastCellParagraph = cellTable.getParagraphArray(0);
				XWPFTable newTable = cellTable.insertNewTbl(lastCellParagraph.getCTP().newCursor());
				
				setNewTableProperty(newTable, nestedFirstRowTable);		
				//newTable.setWidth(cellTable.getWidth());
				
				XWPFTableRow newTableRow = newTable.getRow(0);	
				
				// to create row for 1st time to initialize the creation of table
				newTableRow = newTable.createRow();
				
				// to create cell for row = 0 for 1st time 
				for (int columnIndex = 0; columnIndex < nestedFirstRowTable.getRow(0).getTableCells().size(); columnIndex++) {
					String headerText = nestedFirstRowTable.getRow(0).getCell(columnIndex).getTextRecursively();
					newTableRow.createCell().setText(headerText);
				}
				
				for (int row = 1; row <= newRowSize; row++) {
					newTableRow = newTable.createRow();
					
					for (int cell = 0; cell < newTableRow.getTableCells().size(); cell++) {
						XWPFTableCell tableCell = newTableRow.getTableCells().get(cell);
						
						for (XWPFParagraph cellParagraph : tableCell.getParagraphs()) {							
							insertNewRecordInTableCell(tagsInBetweenCollection, cell, cellParagraph,
									nestedCollectionValues, row, nestedFirstRowTable, tagName);	
						}
					}
				}
			}
		}
		
		else if (OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
			String nonParentNestedCollectionName = DocxUtils.getMapKey(tagName);
			
			if(nestedCollectionDO.getResolutionAttributesMap().containsKey(nonParentNestedCollectionName)) {
				if (nestedCollectionDO.getResolutionAttributesMap().get(nonParentNestedCollectionName) instanceof ArrayList) {
					ListIterator<Object> iterator = ((ArrayList)collectionDO.getResolutionAttributesMap()
							.get(nonParentNestedCollectionName)).listIterator();
					List<Object> newCollectionValues = IteratorUtils.toList(iterator);
					
					int newRowSize = newCollectionValues.size();
					
					XWPFParagraph lastCellParagraph = cellTable.getParagraphArray(0);
					XWPFTable newTable = cellTable.insertNewTbl(lastCellParagraph.getCTP().newCursor());
					
					setNewTableProperty(newTable, nestedFirstRowTable);
					
					XWPFTableRow newTableRow = newTable.getRow(0);	
					
					// to create row for 1st time to initialize the creation of table
					newTableRow = newTable.createRow();
					
					// to create cell for row = 0 for 1st time 
					for (int columnIndex = 0; columnIndex < nestedFirstRowTable.getRow(0).getTableCells().size(); columnIndex++) {
						String firstRowText = nestedFirstRowTable.getRow(0).getCell(columnIndex).getTextRecursively();
						newTableRow.createCell().setText(firstRowText);
					}
					
					for (int row = 1; row <= newRowSize; row++) {
						newTableRow = newTable.createRow();
						
						for (int cell = 0; cell < newTableRow.getTableCells().size(); cell++) {
							XWPFTableCell tableCell = newTableRow.getTableCells().get(cell);
							
							for (XWPFParagraph cellParagraph : tableCell.getParagraphs()) {								
								insertNewRecordInTableCell(tagsInBetweenCollection, cell, cellParagraph,
										newCollectionValues, row, nestedFirstRowTable, tagName);	
							}
						}
					}			
				}
			}					
		}
		
	}

	private void insertNewRecordInTableCell(List<TagDO> tagsInBetweenCollection, int cell, XWPFParagraph cellParagraph,
			List<Object> newCollectionValues, int row, XWPFTable nestedFirstRowTable, String tagName) {

		TagDO nestedTag = tagsInBetweenCollection.get(cell);
		String paragraphText = cellParagraph.getText();
		String nestedTagText = DocxUtils.addTagBracket(nestedTag.getTagText());

		if (COLLECTION_OBJECT_PATTERN.matcher(nestedTagText).matches()) {
			if (DocxUtils.isNullEmpty(paragraphText)) {
				XWPFRun newRun = cellParagraph.createRun();
				String newText = newRun.getText(0);

				if (DocxUtils.isNullEmpty(newText))
					newText = DocxConstants.EMPTY_STRING;

				Object newCollectionValue = newCollectionValues.get(row - 1);
				//tag equals to phone.phoneNo, returns phone
				String collectionName = DocxUtils.getObjectName(nestedTag.getTagText());
				String parentcollectionName= DocxUtils.getObjectField(tagName);

				if (collectionName.equals(parentcollectionName)) {
					//tag equals to user.phones:phone, returns phone
					String objectField = DocxUtils.getObjectField(nestedTag.getTagText());
					String value = DocxUtils.getFieldValue(objectField, newCollectionValue);
					
					String parentParagraphText = nestedFirstRowTable.getRow(1).getCell(cell).getTextRecursively();
					newText = newText.replace(newText, parentParagraphText);
					newRun.setText(newText, 0);
					DocxUtils.replaceTextSegment(cellParagraph, nestedTagText, value);
				}
			}
		}
		
	}
	
	private void insertParagraphInTableCell(TagDO tag, Object collectionValue, XWPFTableCell cellTable) {
		
		String tagText = DocxUtils.addTagBracket(tag.getTagText());
		
		XWPFParagraph newParagraph = cellTable.addParagraph();
		XWPFRun newRun = newParagraph.createRun();
		String newText = newRun.getText(0);
		
		if (DocxUtils.isNullEmpty(newText))
			newText = DocxConstants.EMPTY_STRING;
		
		String objectField = DocxUtils.getObjectField(tag.getTagText());
		String value = DocxUtils.getFieldValue(objectField, collectionValue);
		IBodyElement bodyElement = tag.getTagElement();
		
		if (bodyElement instanceof XWPFParagraph) {
			XWPFParagraph parentParagraph = (XWPFParagraph) bodyElement;
			newText = newText.replace(newText, parentParagraph.getText());
			newRun.setText(newText, 0);
			DocxUtils.replaceTextSegment(newParagraph, tagText, value);
		}	
	}

	private void setNewTableProperty(XWPFTable newTable, XWPFTable nestedFirstRowTable) {
		
		newTable.getCTTbl().addNewTblPr().addNewTblBorders().addNewLeft().setVal(
				org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
		newTable.getCTTbl().getTblPr().getTblBorders().addNewRight().setVal(
				org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
		newTable.getCTTbl().getTblPr().getTblBorders().addNewTop().setVal(
				org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
		newTable.getCTTbl().getTblPr().getTblBorders().addNewBottom().setVal(
				org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
		newTable.getCTTbl().getTblPr().getTblBorders().addNewInsideH().setVal(
				org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
		newTable.getCTTbl().getTblPr().getTblBorders().addNewInsideV().setVal(
				org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
	}

	private void insertNewParagraph(XWPFParagraph paragraph, String value, CollectionDO collectionDO, TagDO tag)
			throws Exception {
	
		if (collectionDO.isElementInTable()) {
			ParentTableDO parentTableDO = collectionDO.getParentTableDO();
			XWPFTable parentTable = parentTableDO.getTable();
			XWPFTableCell parentCellTable = parentTable.getRow(parentTableDO.getRowIndex()).getCell(parentTableDO.getCellIndex());
			
			XWPFParagraph endTagElement = (XWPFParagraph) collectionDO.getEndCollectionElement();
			
			XWPFParagraph newParagraph = parentCellTable.insertNewParagraph(endTagElement.getCTP().newCursor());
			XWPFRun run = newParagraph.createRun();
			IBodyElement bodyElement = tag.getTagElement();
			
			if (bodyElement instanceof XWPFParagraph) {
				XWPFParagraph parentParagraph = (XWPFParagraph) bodyElement;
				run.setText(parentParagraph.getText());
				String textToBeReplaced = DocxUtils.addTagBracket(tag.getTagText());
				DocxUtils.replaceTextSegment(newParagraph, textToBeReplaced, value);
			}
			
		} else {
			XWPFDocument document = paragraph.getDocument();

			XWPFParagraph endTagElement = (XWPFParagraph) collectionDO.getEndCollectionElement();
			XWPFParagraph newParagraph = document.insertNewParagraph(endTagElement.getCTP().newCursor());
			XWPFRun run = newParagraph.createRun();
			IBodyElement bodyElement = tag.getTagElement();
			
			if (bodyElement instanceof XWPFParagraph) {
				XWPFParagraph parentParagraph = (XWPFParagraph) bodyElement;
				run.setText(parentParagraph.getText());
				String textToBeReplaced = DocxUtils.addTagBracket(tag.getTagText());
				DocxUtils.replaceTextSegment(newParagraph, textToBeReplaced, value);
			}			
		}
	}
	
	private XWPFTable insertNewTable(XWPFTable table, IBodyElement endCollectionElement)
			throws Exception {

		XWPFDocument document = table.getBody().getXWPFDocument();

		XWPFParagraph endTagElement = (XWPFParagraph) endCollectionElement;
		XmlCursor cursor = endTagElement.getCTP().newCursor();
		XWPFTable newTable = document.insertNewTbl(cursor);
		
		return newTable;
	}

	private List<TagDO> getTagsFromCollection(IBodyElement elementAfterStartCollection, List<TagDO> tags,
			int startIndex, int endIndex) throws Exception {

		List<IBodyElement> subListBodyElements = elementAfterStartCollection.getBody().getBodyElements();
		
		IBodyElement bodyElem = subListBodyElements.get(startIndex + 1);
		
		while (bodyElem != subListBodyElements.get(endIndex)) {
			tags = getTagsInBetween(bodyElem, tags);
			bodyElem = DocxUtils.getNextElement(bodyElem);
		}

		return tags;
	}

	private List<TagDO> getTagsInBetween(IBodyElement bodyElem, List<TagDO> tags) throws Exception {
		
		if (bodyElem instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			String paragraphText = paragraph.getText();
			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, bodyElem);
			} 	
		}
		
		else if (bodyElem instanceof XWPFTable) {
			XWPFTable table = (XWPFTable) bodyElem;
			for (int row = 0; row < table.getRows().size(); row++) {
				XWPFTableRow tableRow = table.getRows().get(row);
				
				for (int cell = 0; cell < tableRow.getTableCells().size(); cell++) {
					XWPFTableCell tableCell = tableRow.getTableCells().get(cell);
					
					IBodyElement cellBodyElement = tableCell.getBodyElements().get(0);
					while (cellBodyElement != null) {					
						tags = getTagsInBetween(cellBodyElement, tags);
						cellBodyElement = DocxUtils.getNextElement(cellBodyElement);
					}
				}
			}
		}
		
		return tags; 	
	}
	
	public CollectionDO getEndCollection(IBodyElement collectionStartElement, int startCollectionIndex,
			CollectionDO collectionDO, IBodyElement parentCollection) throws Exception {

		List<IBodyElement> subListBodyElements = collectionStartElement.getBody().getBodyElements();
		
		IBodyElement bodyElem = subListBodyElements.get(startCollectionIndex);
		
		while (bodyElem != null) {
			collectionDO = getEndCollectionElement(bodyElem, collectionDO, parentCollection);
			bodyElem = DocxUtils.getNextElement(bodyElem);
		}

		if (collectionDO.getEndCollectionIndex() == null || collectionDO.getEndCollectionElement() == null)
			throw new Exception("No end collection tag found"); 
			
		return collectionDO;
	}
	
	private CollectionDO getEndCollectionElement(IBodyElement bodyElem, CollectionDO collectionDO, IBodyElement parentCollection) throws Exception {
		
		if (bodyElem instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			String paragraphText = paragraph.getText();
			List<TagDO> tags = new ArrayList<>();
			
			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, bodyElem);

				for (TagDO tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
						
						if(DocxUtils.isNullEmpty(collectionDO.getStartCollectionName())) {
							String startCollectionTag = DocxUtils.getTagName(tag.getTagText(),
									DocxConstants.TAG_PREFIX_COLLECTION_START);
							collectionDO.setStartCollectionName(startCollectionTag);
						}	
					}

					else if (COLLECTION_END_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_END_PATTERN_2.matcher(tagText).matches()) {
						String endCollectionTag = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_COLLECTION_END);

						if (endCollectionTag.equals(collectionDO.getStartCollectionName())) {
							collectionDO.setEndCollectionIndex(DocxUtils.getElementIndex(bodyElem));
							
							if (!collectionDO.isNestedCollection()) {
								collectionDO.setEndCollectionElement(bodyElem);
							} else
								collectionDO.setEndCollectionElement(parentCollection);
							break;
						}
					}
				}
			}	
		}
		
		else if (bodyElem instanceof XWPFTable) {
			XWPFTable table = (XWPFTable) bodyElem;
			
			for (int row = 0; row < table.getRows().size(); row++) {
				XWPFTableRow tableRow = table.getRows().get(row);
				
				for (int cell = 0; cell < tableRow.getTableCells().size(); cell++) {
					XWPFTableCell tableCell = tableRow.getTableCells().get(cell);
					
					IBodyElement cellBodyElement = tableCell.getBodyElements().get(0);
					while (cellBodyElement != null) {					
						collectionDO = getEndCollectionElement(cellBodyElement, collectionDO, parentCollection);
						cellBodyElement = DocxUtils.getNextElement(cellBodyElement);
					}
				}
			}
		}

		return collectionDO;
	}

}
