package com.docx.service.tag;

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

import com.docx.service.model.CollectionDO;
import com.docx.service.model.ParentTableDO;
import com.docx.service.model.TagInfo;
import com.docx.service.utils.DocxConstants;
import com.docx.service.utils.DocxUtils;

public class CollectionTagProcessor {
	
	private static final Pattern OBJECT_FIELD_PATTERN_1 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+");
	private static final Pattern OBJECT_FIELD_PATTERN_2 = Pattern.compile("[a-zA-Z]+:[a-zA-Z]+");
	
	private static final Pattern COLLECTION_START_PATTERN_1 = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_2 = Pattern.compile("\\$\\{collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_1 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_2 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");

	public CollectionDO getCollection(CollectionDO collectionDO) {
		collectionDO = getCollectionValues(collectionDO);
		return collectionDO;
	}

	public void newProcess(IBodyElement elementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {
		
		if (elementAfterStartCollection instanceof XWPFParagraph) {
			processCollectionInParagraph(elementAfterStartCollection, collectionDO, resolutionAttributesMap);
		} else if (elementAfterStartCollection instanceof XWPFTable) {
			processCollectionInTable(elementAfterStartCollection, collectionDO, resolutionAttributesMap);
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

	private CollectionDO getCollectionValues(CollectionDO collectionDO) {

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

		return collectionDO;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCollectionInParagraph(IBodyElement elementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {

		List<TagInfo> tags = new ArrayList<>();
		String nonParentNestedCollectionKey = null;

		XWPFParagraph paragraph = (XWPFParagraph) elementAfterStartCollection;
		
		// throw new error here if no end collection index found
		tags = getTagsFromCollection(elementAfterStartCollection, tags,
				collectionDO.getStartCollectionIndex(), collectionDO.getEndCollectionIndex());

		if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {

			for (Object collectionValue : collectionDO.getCollectionValues()) {

				for (TagInfo tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
						//example value listOfUser:user, get field -> user
						String collectionName = DocxUtils.getObjectField(collectionDO.getTagName());
						// example value from user.name -> user
						String objectName = DocxUtils.getObjectName(tag.getTagText());
						if (collectionName.equals(objectName)) {
							// example value from user.name -> name
							String objectField = DocxUtils.getObjectField(tag.getTagText());
							// get the value of the field based on the tag object field
							String value = DocxUtils.getFieldValue(objectField, collectionValue);
							// replace tag text with
							insertNewParagraph(paragraph, value, collectionDO, tag);
						}
					}

					else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
						String tagName = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_COLLECTION_START);

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
								nestedCollectionDO.setResolutionAttributesMap(resolutionAttributesMap);
								nestedCollectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(tag.getTagElement()));
								nestedCollectionDO.setNestedCollection(true);
								nestedCollectionDO = getEndCollection(paragraph, DocxUtils.getElementIndex(tag.getTagElement()),
										nestedCollectionDO, collectionDO.getEndCollectionElement());	
								nestedCollectionDO = getCollection(nestedCollectionDO);
								if (collectionDO.isElementInTable()) {
									nestedCollectionDO.setElementInTable(true);
									nestedCollectionDO.setParentTableDO(collectionDO.getParentTableDO());
								}
								
								IBodyElement nestedCollectionElement = tag.getTagElement();
								
								if (nestedCollectionElement instanceof XWPFParagraph) {
									XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
									//nestedCollectionDO = docxEngine.process(nestedParagraph, tag, newValues, nestedCollectionDO);
									IBodyElement nextElement = DocxUtils.getNextElement(nestedParagraph);
									newProcess(nextElement, nestedCollectionDO, nestedCollectionDO.getResolutionAttributesMap());
								} 
							}
						}

						else if (OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
							nonParentNestedCollectionKey = DocxUtils.getMapKey(tagName); // employees:name -> employees
							
							//get collectionDO start element index
							//get nonparent collection element index
							//traverse all the way back to start index
							//if found other collection start tag, non parent collection has been processed
							//therefore no need to process again
							boolean isProcessed = isCollectionProcessed(tags, tag);
							
							if(!isProcessed) {
								if (resolutionAttributesMap.containsKey(nonParentNestedCollectionKey)) {
									if (resolutionAttributesMap.get(nonParentNestedCollectionKey) instanceof ArrayList) {
										ListIterator<Object> iterator = ((ArrayList)resolutionAttributesMap.get(nonParentNestedCollectionKey)).listIterator();
										List<Object> newCollectionValues = IteratorUtils.toList(iterator);
										
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
										nestedCollectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(tag.getTagElement()));
										nestedCollectionDO.setNestedCollection(true);								
										nestedCollectionDO = getEndCollection(paragraph, DocxUtils.getElementIndex(tag.getTagElement()),
												nestedCollectionDO, collectionDO.getEndCollectionElement());
										if (collectionDO.isElementInTable()) {
											nestedCollectionDO.setElementInTable(true);
											nestedCollectionDO.setParentTableDO(collectionDO.getParentTableDO());
										}
											
										
										IBodyElement nestedCollectionElement = tag.getTagElement();
										
										if (nestedCollectionElement instanceof XWPFParagraph) {
											XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
											//nestedCollectionDO = docxEngine.process(nestedParagraph, tag, newValues, nestedCollectionDO);
											IBodyElement nextElement = DocxUtils.getNextElement(nestedParagraph);
											newProcess(nextElement, nestedCollectionDO, nestedCollectionDO.getResolutionAttributesMap());
										}	
									}								
								}
							}
						}
					}
				}
			}
			
//			if (!DocxUtils.isNullEmpty(nonParentNestedCollectionKey)) {
//				resolutionAttributesMap.remove(nonParentNestedCollectionKey);
//			}		
		}
	}
	
	private boolean isCollectionProcessed(List<TagInfo> tags, TagInfo parentTag)
			throws Exception {
		
		List<TagInfo> subListTags = new ArrayList<>();
		
		//get sublist of tag
		for(TagInfo tag : tags) {
			if(!parentTag.getTagText().equals(tag.getTagText())) {
				subListTags.add(tag);
			} else
				break;
		}
		
		Collections.reverse(subListTags);
		for (TagInfo subListTag : subListTags) {
			String tagText = DocxUtils.addTagBracket(subListTag.getTagText());
			if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
					|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
				return true;
			}
		}
		
		return false;
	}

	private void processNestedCollectionTableInParagraph(XWPFTable tableElementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {
		
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
					List<IBodyElement> firstRowCellBodyElements = parentTable.getRow(1).getCell(cell).getBodyElements();
					
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
					} else {
						// loop each element to get the tags in each cell
						for (IBodyElement firstRowCellBodyElem : firstRowCellBodyElements) {
							List<TagInfo> tags = new ArrayList<>();
							
							if (firstRowCellBodyElem instanceof XWPFParagraph) {
								XWPFParagraph firstRowCellParagraph = (XWPFParagraph) firstRowCellBodyElem;
								String paragraphText = firstRowCellParagraph.getText();

								if (!DocxUtils.isNullEmpty(paragraphText)) {
									tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, firstRowCellBodyElem);
									for (TagInfo tag : tags) {
										processCell(newTableCell, newCellBodyElem, tag, collectionDO,
												row+1, firstRowCellBodyElements, firstRowCellBodyElem, resolutionAttributesMap);
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
	
	private void processCollectionInTable(IBodyElement tableElementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {
		
		XWPFTable table = (XWPFTable) tableElementAfterStartCollection;
		
		// if nested collection in paragraph was a table
		if (collectionDO.isNestedCollection()) {
			processNestedCollectionTableInParagraph(table, collectionDO, resolutionAttributesMap);
		} else {
			// collection in table
			for (int value = 0; value < collectionDO.getCollectionValues().size(); value++) {
				table.createRow();
			}
			
			for (int rowIndex = 2; rowIndex < table.getRows().size(); rowIndex++) {

				XWPFTableRow rowTable = table.getRows().get(rowIndex);

				for (int cellIndex = 0; cellIndex < rowTable.getTableCells().size(); cellIndex++) {
					XWPFTableCell cellTable = rowTable.getTableCells().get(cellIndex);
					// to get the tags in the first row, get the cell's body element
					List<IBodyElement> firstRowCellBodyElements = table.getRow(1).getCell(cellIndex).getBodyElements();
					
					// loop each element to get the tags in each cell
					for (IBodyElement firstRowCellBodyElem : firstRowCellBodyElements) {
						List<TagInfo> tags = new ArrayList<>();
						
						if (firstRowCellBodyElem instanceof XWPFParagraph) {
							XWPFParagraph firstRowCellParagraph = (XWPFParagraph) firstRowCellBodyElem;
							String paragraphText = firstRowCellParagraph.getText();

							if (!DocxUtils.isNullEmpty(paragraphText)) {
								tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, firstRowCellBodyElem);
								
								IBodyElement newCellBodyElem = cellTable.getBodyElements().get(0);
								for (TagInfo tag : tags) {
									processCell(cellTable, newCellBodyElem, tag, collectionDO,
											rowIndex, firstRowCellBodyElements, firstRowCellBodyElem, resolutionAttributesMap);				
								}
							}
						} 
					}
				}
			}

			table.removeRow(1);			
		}
	}

	private void processCell(XWPFTableCell cellTable, IBodyElement newCellBodyElem, TagInfo tag,
			CollectionDO collectionDO, int rowIndex, List<IBodyElement> firstRowCellBodyElements,
			IBodyElement firstRowCellBodyElem, Map<String, Object> resolutionAttributesMap) throws Exception {
		
		String tagText = DocxUtils.addTagBracket(tag.getTagText());
		
		if (newCellBodyElem instanceof XWPFParagraph) {
			
			if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
				
				XWPFParagraph paragraph = (XWPFParagraph) newCellBodyElem;
				String text = paragraph.getText();
				if (DocxUtils.isNullEmpty(text)) {
					if (firstRowCellBodyElem instanceof XWPFParagraph) {
						XWPFParagraph parentParagraph = (XWPFParagraph) firstRowCellBodyElem;
						String parentParagraphText = parentParagraph.getText();
						text = text.replace(text, parentParagraphText);
						XWPFRun run = paragraph.createRun();
						run.setText(parentParagraphText, 0);
						
						//get the collection value
						Object collectionValue = collectionDO.getCollectionValues().get(rowIndex - 2);
						String collectionObject = DocxUtils.getObjectName(tag.getTagText()); // tag text = user.name, get user
						String collectionField= DocxUtils.getObjectField(collectionDO.getTagName()); // tag name = user.name, get user
						
						if (collectionObject.equals(collectionField)) {
							String objectField = DocxUtils.getObjectField(tag.getTagText());
							String value = DocxUtils.getFieldValue(objectField, collectionValue);
							
							DocxUtils.replaceTextSegment(paragraph, tagText, value);
						}	
					}
				} else {
					XWPFParagraph newParagraph = cellTable.addParagraph();
					XWPFRun newRun = newParagraph.createRun();
					String newText = newRun.getText(0);
					
					if (DocxUtils.isNullEmpty(newText))
						newText = DocxConstants.EMPTY_STRING;
					
					//get the collection value
					Object collectionValue = collectionDO.getCollectionValues().get(rowIndex - 2);
					String collectionObject = DocxUtils.getObjectName(tag.getTagText()); // tag text = user.name, get user
					String collectionField= DocxUtils.getObjectField(collectionDO.getTagName()); // tag name = user.name, get user
					
					if (collectionObject.equals(collectionField)) {
						String objectField = DocxUtils.getObjectField(tag.getTagText());
						String value = DocxUtils.getFieldValue(objectField, collectionValue);
						
						if (firstRowCellBodyElem instanceof XWPFParagraph) {
							XWPFParagraph parentParagraph = (XWPFParagraph) firstRowCellBodyElem;
							String parentParagraphText = parentParagraph.getText();
							newText = newText.replace(newText, parentParagraphText);
							newRun.setText(newText, 0);
							DocxUtils.replaceTextSegment(newParagraph, tagText, value);
						}
					}
				}
			}
			
			else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches() 
					|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
				
				String tagName = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_COLLECTION_START);
				String nestedCollectionKey = DocxUtils.getMapKey(tagName);
				
				CollectionDO nestedCollectionDO = new CollectionDO();
				nestedCollectionDO.setTagName(tagName);			
				nestedCollectionDO.setMapKey(nestedCollectionKey);
				nestedCollectionDO.setResolutionAttributesMap(collectionDO.getResolutionAttributesMap());
				nestedCollectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(newCellBodyElem));
				nestedCollectionDO.setNestedCollection(true);
				nestedCollectionDO = getEndCollection(firstRowCellBodyElem, DocxUtils.getElementIndex(tag.getTagElement()),
						nestedCollectionDO, collectionDO.getEndCollectionElement());
	
				List<TagInfo> tagsInBetweenCollection = new ArrayList<>();
				tagsInBetweenCollection = getTagsFromCollection(firstRowCellBodyElem, tagsInBetweenCollection,
						DocxUtils.getElementIndex(newCellBodyElem), nestedCollectionDO.getEndCollectionIndex());
						
				Object collectionValue = collectionDO.getCollectionValues().get(rowIndex - 2);
				
				IBodyElement firstRowCellNextElement = DocxUtils.getNextElement(firstRowCellBodyElem);

				if (firstRowCellNextElement instanceof XWPFParagraph) {
					processCollectionParagraphInCell(tagName, collectionDO, collectionValue, cellTable,
							tagsInBetweenCollection, nestedCollectionDO);
				}
				
				else if (firstRowCellNextElement instanceof XWPFTable) {			
					processCollectionTableInCell(firstRowCellNextElement, tagName, collectionDO, collectionValue, cellTable,
							tagsInBetweenCollection, nestedCollectionDO);
				}	
			}
		}
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCollectionParagraphInCell(String tagName, CollectionDO collectionDO,
			Object collectionValue, XWPFTableCell cellTable, List<TagInfo> tagsInBetweenCollection,
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

					for (TagInfo nestedTag : tagsInBetweenCollection) {
						
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

						for (TagInfo nestedTag : tagsInBetweenCollection) {
							
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
			Object collectionValue, XWPFTableCell cellTable, List<TagInfo> tagsInBetweenCollection,
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

	private void insertNewRecordInTableCell(List<TagInfo> tagsInBetweenCollection, int cell, XWPFParagraph cellParagraph,
			List<Object> newCollectionValues, int row, XWPFTable nestedFirstRowTable, String tagName) {

		TagInfo nestedTag = tagsInBetweenCollection.get(cell);
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
				String collectionField= DocxUtils.getObjectField(tagName);

				if (collectionName.equals(collectionField)) {
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
	
	private void insertParagraphInTableCell(TagInfo tag, Object collectionValue, XWPFTableCell cellTable) {
		
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
		
		//newTable.setWidth(nestedFirstRowTable.getWidth());
	}

	private void insertNewParagraph(XWPFParagraph paragraph, String value, CollectionDO collectionDO, TagInfo tag)
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

	private List<TagInfo> getTagsFromCollection(IBodyElement elementAfterStartCollection, List<TagInfo> tags,
			int startIndex, int endIndex) throws Exception {

		List<IBodyElement> subListBodyElements = elementAfterStartCollection.getBody().getBodyElements();
		
		IBodyElement bodyElem = subListBodyElements.get(startIndex + 1);
		
//		if(endIndex > subListBodyElements.size())
//			endIndex = subListBodyElements.size()-1;
		while (bodyElem != subListBodyElements.get(endIndex)) {
			tags = getTagsInBetween(bodyElem, tags);
			bodyElem = DocxUtils.getNextElement(bodyElem);
		}

		return tags;
	}

	private List<TagInfo> getTagsInBetween(IBodyElement bodyElem, List<TagInfo> tags) throws Exception {
		
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
			List<TagInfo> tags = new ArrayList<>();
			
			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, bodyElem);

				for (TagInfo tag : tags) {
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
