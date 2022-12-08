package com.template.engine.tag;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections.IteratorUtils;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import com.template.engine.DocxEngine;
import com.template.engine.model.CollectionDO;
import com.template.engine.model.ParentTableDO;
import com.template.engine.model.TagInfo;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class CollectionTagProcessor {

	private static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_OBJECT_FIELD_PATTERN_1 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+");
	private static final Pattern COLLECTION_OBJECT_FIELD_PATTERN_2 = Pattern.compile("[a-zA-Z]+:[a-zA-Z]+");
	private static final Pattern COLLECTION_OBJECT_FIELD_PATTERN_3 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+");
	private static final Pattern COLLECTION_START_PATTERN_1 = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_2 = Pattern.compile("\\$\\{collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_1 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_2 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");

	public CollectionDO getCollection(CollectionDO collectionDO) {
		collectionDO = getCollectionValues(collectionDO);
		return collectionDO;
	}

	public void newProcess(IBodyElement elementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {
		
		if (elementAfterStartCollection instanceof XWPFParagraph) {
			processCollectionInParagraph((XWPFParagraph) elementAfterStartCollection, collectionDO, resolutionAttributesMap);
		} else if (elementAfterStartCollection instanceof XWPFTable) {
			processCollectionInTable((XWPFTable) elementAfterStartCollection, collectionDO);
		}
	}

	public String getFirstParameter(String tag) {

		String firstParameter = null;
		int indexOfColon = tag.indexOf(":", 0);
		if (indexOfColon > 0) {
			firstParameter = tag.substring(0, indexOfColon);
		}

		return firstParameter;
	}

	public String getSecondParameter(String tag) {

		String secondParameter = null;
		int indexOfColon = tag.indexOf(":", 0);
		if (indexOfColon > 0) {
			secondParameter = tag.substring(indexOfColon + 1, tag.length());
			;
		}

		return secondParameter;
	}

	public String getFirstParameterTypeTwo(String tag) {

		String firstParameter = null;
		int indexOfDot = tag.indexOf(".", 0);
		if (indexOfDot > 0) {
			firstParameter = tag.substring(0, indexOfDot);
		}
		return firstParameter;
	}

	public String getSecondParameterTypeTwo(String tag) {

		String secondParameter = null;
		int indexOfDot = tag.indexOf(".", 0);
		if (indexOfDot > 0) {
			secondParameter = tag.substring(indexOfDot + 1, tag.length());
		}

		return secondParameter;
	}

	public String getClassName(String tag) {

		String className = null;
		if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(tag).matches()) {
			className = getFirstParameter(tag);
			className = getFirstParameterTypeTwo(className);
		} else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(tag).matches()) {
			className = getFirstParameter(tag);
		} else if (COLLECTION_OBJECT_FIELD_PATTERN_3.matcher(tag).matches()) {
			className = getFirstParameterTypeTwo(tag);
		}
		return className;
	}

	public String getClassField(String tag) {

		String classField = null;
		if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(tag).matches()
				|| COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(tag).matches()) {
			classField = getSecondParameter(tag);
		} else if (COLLECTION_OBJECT_FIELD_PATTERN_3.matcher(tag).matches()) {
			classField = getSecondParameterTypeTwo(tag);
		}
		return classField;
	}

	private List<Object> getCollectionValues(String objectName, Object mapValue) {

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

		// check type pattern
		if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(collectionDO.getTagName()).matches()) {			
			String collectionObjectName = getClassField(collectionDO.getMapKey());		
			List<Object> collectionValues = getCollectionValues(collectionObjectName,
					collectionDO.getResolutionAttributesMap().get(collectionDO.getMapKey()));		
			collectionDO.setCollectionValues(collectionValues);
		}

		else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(collectionDO.getTagName()).matches()) {
			String collectionField = getSecondParameter(collectionDO.getTagName());
			List<Object> collectionValues = getCollectionValues(collectionField,
					collectionDO.getResolutionAttributesMap().get(collectionDO.getMapKey()));
			collectionDO.setCollectionValues(collectionValues);
		}

		return collectionDO;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCollectionInParagraph(XWPFParagraph elementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {

		List<TagInfo> tags = new ArrayList<>();
		String nonParentNestedCollectionName = null;

		XWPFParagraph paragraph = (XWPFParagraph) elementAfterStartCollection;
				
		tags = getTagsFromCollection(elementAfterStartCollection, tags,
				collectionDO.getStartCollectionIndex() + 1, collectionDO.getEndCollectionIndex());

		if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {

			for (Object collectionValue : collectionDO.getCollectionValues()) {

				for (TagInfo tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
						//example value listOfUser:user, get field -> user
						String collectionName = getClassField(collectionDO.getTagName());
						// example value from user.name -> user
						String objectName = getClassName(tag.getTagText());
						if (collectionName.equals(objectName)) {
							// example value from user.name -> name
							String objectField = getClassField(tag.getTagText());
							// get the value of the field based on the tag object field
							String value = DocxUtils.getFieldValue(objectField, collectionValue);
							boolean isLastTag = tag == tags.get(tags.size() - 1);
							// replace tag text with
							insertNewParagraph(paragraph, value, isLastTag, collectionDO);
						}
					}

					else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
						String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_START);
						DocxEngine docxEngine = new DocxEngine();

						if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
							// returns example value user.phones:phone -> user.phones
							String nestedCollectionKey = getFirstParameter(tagName);
							// returns user from user.phones
							String nestedCollectionName = getClassName(nestedCollectionKey);
							//example value listOfUser:user, get field -> user
							String parentCollectionName = getClassField(collectionDO.getTagName());
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
								} else
									nestedCollectionDO.setLastCollectionValue(false);
								nestedCollectionDO.setTagName(tagName);
								nestedCollectionDO.setMapKey(nestedCollectionKey);
								nestedCollectionDO.setResolutionAttributesMap(newValues);
								nestedCollectionDO.setStartCollectionIndex(tag.getTagElementIndex());
								nestedCollectionDO.setNestedCollection(true);
								nestedCollectionDO = getEndCollectionIndex(paragraph, tag.getTagElementIndex(),
										nestedCollectionDO, collectionDO.getEndCollectionElement());	
								nestedCollectionDO = getCollection(nestedCollectionDO);
								
								IBodyElement nestedCollectionElement = elementAfterStartCollection.getBody()
										.getBodyElements().get(tag.getTagElementIndex());
								
								if (nestedCollectionElement instanceof XWPFParagraph) {
									XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
									docxEngine.process(nestedParagraph, tag, newValues, nestedCollectionDO);
								} 
							}
						}

						else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
							nonParentNestedCollectionName = getFirstParameter(tagName); // employees:name -> employees
							
							if (resolutionAttributesMap.containsKey(nonParentNestedCollectionName)) {
								if (resolutionAttributesMap.get(nonParentNestedCollectionName) instanceof ArrayList) {
									ListIterator<Object> iterator = ((ArrayList)resolutionAttributesMap.get(nonParentNestedCollectionName)).listIterator();
									List<Object> newCollectionValues = IteratorUtils.toList(iterator);
									
									Map<String, Object> newValues = new HashMap<String, Object>();
									newValues.put(nonParentNestedCollectionName, newCollectionValues);
									
									CollectionDO nestedCollectionDO = new CollectionDO();		
									boolean isLastCollectionValue = collectionValue == collectionDO
											.getCollectionValues().get(collectionDO.getCollectionValues().size() - 1);
									if (isLastCollectionValue) {
										nestedCollectionDO.setLastCollectionValue(true);
									} else
										nestedCollectionDO.setLastCollectionValue(false);						
									nestedCollectionDO.setCollectionValues(newCollectionValues);
									nestedCollectionDO.setTagName(tagName);
									nestedCollectionDO.setMapKey(nonParentNestedCollectionName);
									nestedCollectionDO.setResolutionAttributesMap(newValues); //set resolution attribute
									nestedCollectionDO.setStartCollectionIndex(tag.getTagElementIndex());
									nestedCollectionDO.setNestedCollection(true);								
									nestedCollectionDO = getEndCollectionIndex(paragraph, tag.getTagElementIndex(),
											nestedCollectionDO, collectionDO.getEndCollectionElement());	
									
									IBodyElement nestedCollectionElement = elementAfterStartCollection.getBody()
											.getBodyElements().get(tag.getTagElementIndex());
									
									if (nestedCollectionElement instanceof XWPFParagraph) {
										XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
										docxEngine.process(nestedParagraph, tag, newValues, nestedCollectionDO);									
									}									
								}								
							}
						}
					}
				}
			}
			
			if (!DocxUtils.isNullEmpty(nonParentNestedCollectionName)) {
				resolutionAttributesMap.remove(nonParentNestedCollectionName);
			}
		}

	}
	
	private void processCollectionInTableNestedParagraph(XWPFTable tableElementAfterStartCollection, CollectionDO collectionDO) throws Exception {
		
		XWPFTable parentTable = (XWPFTable) tableElementAfterStartCollection;
		
		IBodyElement endCollectionElement = collectionDO.getEndCollectionElement();
		XWPFTable newTable = insertNewTable(parentTable, endCollectionElement);
		
		if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {
			
			int newCellSize = collectionDO.getCollectionValues().get(0).getClass().getDeclaredFields().length;
			int newRowSize = collectionDO.getCollectionValues().size();
			
			for (int cell = 1; cell < newCellSize; cell++) {
				newTable.getRow(0).createCell();
			}
			
			for (int row = 0; row < newRowSize; row++) {
				newTable.createRow();
			}

			for (int row = 0; row < newTable.getRows().size(); row++) {
				XWPFTableRow newTableRow = newTable.getRows().get(row);
				
				for (int cell = 0; cell < newTableRow.getTableCells().size(); cell++) {
					XWPFTableCell newTableCell = newTableRow.getTableCells().get(cell);
					// first row for each column text of parent table
					String firstRowText = parentTable.getRow(1).getCell(cell).getTextRecursively();
					List<TagInfo> tags = new ArrayList<>();
					tags = DocxUtils.getTagsByElement(firstRowText, 0, tags);
					
					for (XWPFParagraph paragraph : newTableCell.getParagraphs()) {
						
						for (TagInfo tag : tags) {
							//set headers of the table
							if (row == 0) {
								XWPFRun run = paragraph.createRun();
								String text = run.getText(0);
								
								if (DocxUtils.isNullEmpty(text))
									text = DocxConstants.EMPTY_STRING;
								
								String headerText = parentTable.getRow(0).getCell(cell).getTextRecursively(); // set in value to be replaced
								text = text.replace(text, headerText);
								run.setText(text, 0);
							} 
							
							else {
								String paragraphText = paragraph.getText();
								String tagText = DocxUtils.addTagBracket(tag.getTagText());

								if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
									if (DocxUtils.isNullEmpty(paragraphText)) {
										XWPFRun run = paragraph.createRun();
										String text = run.getText(0);

										if (DocxUtils.isNullEmpty(text))
											text = DocxConstants.EMPTY_STRING;

										Object collectionValue = collectionDO.getCollectionValues().get(row - 1);
										//tag equals to phone.phoneNo, returns phone
										String objectName = getClassName(tag.getTagText());
										String collectionName= getClassField(collectionDO.getTagName());

										if (collectionName.equals(objectName)) {
											//tag equals to user.phones:phone, returns phone
											String objectField = getClassField(tag.getTagText());
											String value = DocxUtils.getFieldValue(objectField, collectionValue);
											text = text.replace(text, value);
											run.setText(text, 0);
										}
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
	
	private void processCollectionInTable(XWPFTable tableElementAfterStartCollection, CollectionDO collectionDO)
			throws Exception {
		
		// if nested collection in paragraph was a table
		if (collectionDO.isNestedCollection()) {
			processCollectionInTableNestedParagraph(tableElementAfterStartCollection, collectionDO);
		} else {
			// collection in table
			for (int value = 0; value < collectionDO.getCollectionValues().size(); value++) {
				tableElementAfterStartCollection.createRow();
			}
			
			for (int rowIndex = 2; rowIndex < tableElementAfterStartCollection.getRows().size(); rowIndex++) {

				XWPFTableRow rowTable = tableElementAfterStartCollection.getRows().get(rowIndex);

				for (int cellIndex = 0; cellIndex < rowTable.getTableCells().size(); cellIndex++) {
					XWPFTableCell cellTable = rowTable.getTableCells().get(cellIndex);
					// to get the tags in the first row, get the cell's body element
					List<IBodyElement> firstRowCellBodyElements = tableElementAfterStartCollection.getRow(1).getCell(cellIndex).getBodyElements();
					
					// loop each element to get the tags in each cell
					for (IBodyElement firstRowCellBodyElem : firstRowCellBodyElements) {
						List<TagInfo> tags = new ArrayList<>();
						if (firstRowCellBodyElem instanceof XWPFParagraph) {
							XWPFParagraph firstRowCellParagraph = (XWPFParagraph) firstRowCellBodyElem;
							String paragraphText = firstRowCellParagraph.getText();

							if (!DocxUtils.isNullEmpty(paragraphText)) {
								tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);
								
								IBodyElement newCellBodyElem = cellTable.getBodyElements().get(0);
								for (TagInfo tag : tags) {
									processCell(cellTable, newCellBodyElem, tag, collectionDO,
											rowIndex, firstRowCellBodyElements, firstRowCellBodyElem);				
								}
							}
						} 
					}
				}
			}

			tableElementAfterStartCollection.removeRow(1);			
		}
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void processCell(XWPFTableCell cellTable, IBodyElement newCellBodyElem, TagInfo tag, CollectionDO collectionDO,
			int rowIndex, List<IBodyElement> firstRowCellBodyElements, IBodyElement firstRowCellBodyElem)
					throws Exception {
		
		String tagText = DocxUtils.addTagBracket(tag.getTagText());
		
		if (newCellBodyElem instanceof XWPFParagraph) {
			if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
				
				XWPFParagraph paragraph = cellTable.addParagraph();
				XWPFRun run = paragraph.createRun();
				String text = run.getText(0);
				
				if (DocxUtils.isNullEmpty(text))
					text = DocxConstants.EMPTY_STRING;
				
				//get the collection value
				Object collectionValue = collectionDO.getCollectionValues().get(rowIndex - 2);
				String collectionObject = getClassName(tag.getTagText()); // tag text = user.name, get user
				String collectionField= getClassField(collectionDO.getTagName()); // tag name = user.name, get user
				
				if (collectionObject.equals(collectionField)) {
					String objectField = getClassField(tag.getTagText());
					String value = DocxUtils.getFieldValue(objectField, collectionValue);
					text = text.replace(text, value);
					run.setText(text, 0);				
				}
			}
			
			else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches() 
					|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
				
				String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_START);
				
				int tagIndex = DocxUtils.getElementIndex(newCellBodyElem);
				List<TagInfo> tagsInBetweenCollection = getTagsFromFirstRow(firstRowCellBodyElements, tagIndex, tagName);
				
				Object collectionValue = collectionDO.getCollectionValues().get(rowIndex - 2);
				
				IBodyElement firstRowCellNextElement = DocxUtils.getNextElement(firstRowCellBodyElem);

				if (firstRowCellNextElement instanceof XWPFParagraph) {
					
					if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
						// returns example value user.phones:phone -> user.phones
						String nestedCollectionName = getFirstParameter(tagName);
						// returns user from user.phones
						String nestedCollectionObject = getClassName(nestedCollectionName);
						//example value listOfUser:user, get field -> user
						String nestedCollectionParentObject = getClassField(collectionDO.getTagName());
						
						if (nestedCollectionObject.equals(nestedCollectionParentObject)) {
							String nestedCollectionObjectField = getClassField(nestedCollectionName); // phones
							List<Object> nestedCollectionValues = getCollectionValues(nestedCollectionObjectField, collectionValue);
							
							for (Object nestedCollectionValue : nestedCollectionValues) {

								for (TagInfo nestedTag : tagsInBetweenCollection) {
									
									XWPFParagraph newParagraph = cellTable.addParagraph();
									XWPFRun newRun = newParagraph.createRun();
									String newText = newRun.getText(0);
									
									if (DocxUtils.isNullEmpty(newText))
										newText = DocxConstants.EMPTY_STRING;
									
									String objectField = getClassField(nestedTag.getTagText());
									String value = DocxUtils.getFieldValue(objectField, nestedCollectionValue);
									newText = newText.replace(newText, value);
									newRun.setText(newText, 0);
								}
							}
						}
					}
					
					else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
						String nonParentNestedCollectionName = getFirstParameter(tagName); 
						
						if(collectionDO.getResolutionAttributesMap().containsKey(nonParentNestedCollectionName)) {
							
							if (collectionDO.getResolutionAttributesMap().get(nonParentNestedCollectionName) instanceof ArrayList) {
								ListIterator<Object> iterator = ((ArrayList)collectionDO.getResolutionAttributesMap()
										.get(nonParentNestedCollectionName)).listIterator();
								List<Object> newCollectionValues = IteratorUtils.toList(iterator);
								
								for (Object nestedCollectionValue : newCollectionValues) {

									for (TagInfo nestedTag : tagsInBetweenCollection) {
										
										XWPFParagraph newParagraph = cellTable.addParagraph();
										XWPFRun newRun = newParagraph.createRun();
										String newText = newRun.getText(0);
										
										if (DocxUtils.isNullEmpty(newText))
											newText = DocxConstants.EMPTY_STRING;
										
										String objectField = getClassField(nestedTag.getTagText());
										String value = DocxUtils.getFieldValue(objectField, nestedCollectionValue);
										newText = newText.replace(newText, value);
										newRun.setText(newText, 0);
									}
								}															
							}
						}
					}
				}
				
				else if (firstRowCellNextElement instanceof XWPFTable) {
					
					XWPFTable nestedFirstRowTable = (XWPFTable) firstRowCellNextElement;
					
					if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
						// returns example value user.phones:phone -> user.phones
						String nestedCollectionName = getFirstParameter(tagName);
						// returns user from user.phones
						String nestedCollectionObject = getClassName(nestedCollectionName);
						//example value listOfUser:user, get field -> user
						String nestedCollectionParentObject = getClassField(collectionDO.getTagName());
						
						if (nestedCollectionObject.equals(nestedCollectionParentObject)) {
							String nestedCollectionObjectField = getClassField(nestedCollectionName); // phones
							List<Object> nestedCollectionValues = getCollectionValues(nestedCollectionObjectField, collectionValue);
							
							int newColumnSize = nestedCollectionValues.get(0).getClass().getDeclaredFields().length;
							int newRowSize = nestedCollectionValues.size();
							
							XWPFParagraph lastCellParagraph = cellTable.getParagraphArray(0);
							XWPFTable newTable = cellTable.insertNewTbl(lastCellParagraph.getCTP().newCursor());
							
							setNewTableProperty(newTable);
							
							XWPFTableRow newTableRow = newTable.getRow(0);	
							
							// to create row for 1st time to initialize the creation of table
							newTableRow = newTable.createRow();
							
							// to create cell for row = 0 for 1st time 
							for (int columnIndex = 0; columnIndex < newColumnSize; columnIndex++) {
								String firstRowText = nestedFirstRowTable.getRow(0).getCell(columnIndex).getTextRecursively();
								newTableRow.createCell().setText(firstRowText);
							}
							
							for (int row = 1; row <= newRowSize; row++) {
								newTableRow = newTable.createRow();
								
								for (int cell = 0; cell < newTableRow.getTableCells().size(); cell++) {
									XWPFTableCell tableCell = newTableRow.getTableCells().get(cell);
									
									for (XWPFParagraph cellParagraph : tableCell.getParagraphs()) {
										
										TagInfo nestedTag = tagsInBetweenCollection.get(cell);
										String paragraphText = cellParagraph.getText();
										String nestedTagText = DocxUtils.addTagBracket(nestedTag.getTagText());

										if (COLLECTION_OBJECT_PATTERN.matcher(nestedTagText).matches()) {
											if (DocxUtils.isNullEmpty(paragraphText)) {
												XWPFRun run = cellParagraph.createRun();
												String text = run.getText(0);

												if (DocxUtils.isNullEmpty(text))
													text = DocxConstants.EMPTY_STRING;

												Object nestedCollectionValue = nestedCollectionValues.get(row - 1);
												//tag equals to phone.phoneNo, returns phone
												String collectionName = getClassName(nestedTag.getTagText());
												String collectionField= getClassField(tagName);

												if (collectionName.equals(collectionField)) {
													//tag equals to user.phones:phone, returns phone
													String objectField = getClassField(nestedTag.getTagText());
													String value = DocxUtils.getFieldValue(objectField, nestedCollectionValue);
													text = text.replace(text, value);
													run.setText(text, 0);
												}
											}
										}	
									}
								}
							}
						}
					}
					
					else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
						String nonParentNestedCollectionName = getFirstParameter(tagName);
						
						if(collectionDO.getResolutionAttributesMap().containsKey(nonParentNestedCollectionName)) {
							if (collectionDO.getResolutionAttributesMap().get(nonParentNestedCollectionName) instanceof ArrayList) {
								ListIterator<Object> iterator = ((ArrayList)collectionDO.getResolutionAttributesMap()
										.get(nonParentNestedCollectionName)).listIterator();
								List<Object> newCollectionValues = IteratorUtils.toList(iterator);
								
								int newColumnSize = newCollectionValues.get(0).getClass().getDeclaredFields().length;
								int newRowSize = newCollectionValues.size();
								
								XWPFParagraph lastCellParagraph = cellTable.getParagraphArray(0);
								XWPFTable newTable = cellTable.insertNewTbl(lastCellParagraph.getCTP().newCursor());
								
								setNewTableProperty(newTable);
								
								XWPFTableRow newTableRow = newTable.getRow(0);	
								
								// to create row for 1st time to initialize the creation of table
								newTableRow = newTable.createRow();
								
								// to create cell for row = 0 for 1st time 
								for (int columnIndex = 0; columnIndex < newColumnSize; columnIndex++) {
									String firstRowText = nestedFirstRowTable.getRow(0).getCell(columnIndex).getTextRecursively();
									newTableRow.createCell().setText(firstRowText);
								}
								
								for (int row = 1; row <= newRowSize; row++) {
									newTableRow = newTable.createRow();
									
									for (int cell = 0; cell < newTableRow.getTableCells().size(); cell++) {
										XWPFTableCell tableCell = newTableRow.getTableCells().get(cell);
										
										for (XWPFParagraph cellParagraph : tableCell.getParagraphs()) {
											
											TagInfo nestedTag = tagsInBetweenCollection.get(cell);
											String paragraphText = cellParagraph.getText();
											String nestedTagText = DocxUtils.addTagBracket(nestedTag.getTagText());

											if (COLLECTION_OBJECT_PATTERN.matcher(nestedTagText).matches()) {
												if (DocxUtils.isNullEmpty(paragraphText)) {
													XWPFRun run = cellParagraph.createRun();
													String text = run.getText(0);

													if (DocxUtils.isNullEmpty(text))
														text = DocxConstants.EMPTY_STRING;

													Object newCollectionValue = newCollectionValues.get(row - 1);
													//tag equals to phone.phoneNo, returns phone
													String collectionName = getClassName(nestedTag.getTagText());
													String collectionField= getClassField(tagName);

													if (collectionName.equals(collectionField)) {
														//tag equals to user.phones:phone, returns phone
														String objectField = getClassField(nestedTag.getTagText());
														String value = DocxUtils.getFieldValue(objectField, newCollectionValue);
														text = text.replace(text, value);
														run.setText(text, 0);
													}
												}
											}	
										}
									}
								}			
							}
						}					
					}
				}	
			}
		}
		
	}
	
	private void setNewTableProperty(XWPFTable newTable) {
		
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

	private List<TagInfo> getTagsFromFirstRow(List<IBodyElement> firstRowCellBodyElements, int tagStartIndex, String startCollectionTag)
			throws Exception {
		
		String endCollectionTag = null;
		List<TagInfo> tagsBetweenCollection = new ArrayList<>();
		
		for (int index = tagStartIndex + 1; index < firstRowCellBodyElements.size(); index++) {
			IBodyElement firstRowCellBodyElem = firstRowCellBodyElements.get(index);
			
			if (firstRowCellBodyElem instanceof XWPFParagraph) {
				XWPFParagraph paragraph = (XWPFParagraph) firstRowCellBodyElem;
				String paragraphText = paragraph.getText();
				List<TagInfo> tags = new ArrayList<>();
				
				if (!DocxUtils.isNullEmpty(paragraphText)) {
					tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);
					
					for (TagInfo tag : tags) {
						String tagText = DocxUtils.addTagBracket(tag.getTagText());
						
						if (paragraphText.contains(tagText)) {
							
							if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches() ) {
								tagsBetweenCollection.add(tag);
							}
							
							else if (COLLECTION_END_PATTERN_1.matcher(tagText).matches()
									|| COLLECTION_END_PATTERN_2.matcher(tagText).matches()) {
								endCollectionTag = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_END);
								
								if (endCollectionTag.equals(startCollectionTag)) 
									break;
							}
						}
					}	
				}				
			} 
			
			else if (firstRowCellBodyElem instanceof XWPFTable ) {
				XWPFTable table = (XWPFTable) firstRowCellBodyElem;

				for (XWPFTableRow row : table.getRows()) {
					for (XWPFTableCell cell : row.getTableCells()) {
						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							String paragraphText = paragraph.getText();
							List<TagInfo> tags = new ArrayList<>();
							if (!DocxUtils.isNullEmpty(paragraphText)) {
								tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);
								
								for (TagInfo tag : tags) {
									String tagText = DocxUtils.addTagBracket(tag.getTagText());
									
									if (paragraphText.contains(tagText)) {
										
										if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches() ) {
											tagsBetweenCollection.add(tag);
										}
										
										else if (COLLECTION_END_PATTERN_1.matcher(tagText).matches()
												|| COLLECTION_END_PATTERN_2.matcher(tagText).matches()) {
											endCollectionTag = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_END);
											
											if (endCollectionTag.equals(startCollectionTag)) 
												break;
										}
									}
								}	
							}
						}
					}
				}
			}
		}
		
		return tagsBetweenCollection;
		
	}

	private void insertNewParagraph(XWPFParagraph paragraph, String value, boolean isLastTag, CollectionDO collectionDO)
			throws Exception {
	
		if (collectionDO.isElementInTable()) {
			ParentTableDO parentTableDO = collectionDO.getParentTableDO();
			XWPFTable parentTable = parentTableDO.getTable();
			XWPFTableCell parentCellTable = parentTable.getRow(parentTableDO.getRowIndex()).getCell(parentTableDO.getCellIndex());
			
			XWPFParagraph endTagElement = (XWPFParagraph) collectionDO.getEndCollectionElement();
			
			XWPFParagraph newParagraph = parentCellTable.insertNewParagraph(endTagElement.getCTP().newCursor());
			XWPFRun run = newParagraph.createRun();
			run.setText(value);
			
		} else {
			XWPFDocument document = paragraph.getDocument();

			XWPFParagraph endTagElement = (XWPFParagraph) collectionDO.getEndCollectionElement();
			XWPFParagraph newParagraph = document.insertNewParagraph(endTagElement.getCTP().newCursor());
			XWPFRun run = newParagraph.createRun();
			run.setText(value);
		}
		
		// isLastTag TODO
	}
	
	private XWPFTable insertNewTable(XWPFTable table, IBodyElement endCollectionElement)
			throws Exception {

		XWPFDocument document = table.getBody().getXWPFDocument();

		XWPFParagraph endTagElement = (XWPFParagraph) endCollectionElement;
		XmlCursor cursor = endTagElement.getCTP().newCursor();
		XWPFTable newTable = document.insertNewTbl(cursor);
		
		return newTable;
	}

	private List<TagInfo> getTagsFromCollection(XWPFParagraph elementAfterStartCollection, List<TagInfo> tags, int startIndex, int endIndex) throws Exception {

		List<IBodyElement> subListBodyElements = elementAfterStartCollection.getBody().getBodyElements();
		for (int tagElementIndex = startIndex; tagElementIndex < endIndex; tagElementIndex++) {
			IBodyElement bodyElem = subListBodyElements.get(tagElementIndex);

			if (bodyElem instanceof XWPFParagraph) {
				XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
				String paragraphText = paragraph.getText();

				if (!DocxUtils.isNullEmpty(paragraphText)) {
					tags = getTagsInBetweenCollection(paragraphText, 0, tagElementIndex, tags);
				}
			}
		}

		return tags;
	}

	private List<TagInfo> getTagsInBetweenCollection(String elementText, int tagStartOffset, int tagElementIndex,
			List<TagInfo> tags) throws Exception {

		tagStartOffset = elementText.indexOf(DocxConstants.DEFAULT_TAG_START, tagStartOffset);
		if (tagStartOffset >= 0) {
			int tagEndOffset = elementText.indexOf(DocxConstants.DEFAULT_TAG_END, tagStartOffset);

			if (tagEndOffset < 0) {
				throw new Exception("No closing tag found for line " + elementText);
			}

			String tagText = elementText.substring(tagStartOffset + DocxConstants.DEFAULT_TAG_START.length(),
					tagEndOffset);
			boolean hasClosingSlash = tagText.startsWith("/");

			TagInfo tagInfo = new TagInfo(tagText, tagEndOffset, hasClosingSlash, tagElementIndex);
			tags.add(tagInfo);

			// recursive to proceed to get other tags in the same text
			getTagsInBetweenCollection(elementText, tagEndOffset, tagElementIndex, tags);
		}

		return tags;
	}

	public CollectionDO getEndCollectionIndex(IBodyElement elementAfterStartCollection, int startCollectionIndex,
			CollectionDO collectionDO, IBodyElement parentCollection) throws Exception {

		List<IBodyElement> bodyElements = elementAfterStartCollection.getBody().getBodyElements();
		
		String startCollectionTag = null;
		String endCollectionTag = null;

		for (int i = startCollectionIndex; i < bodyElements.size(); i++) {
			IBodyElement bodyElem = bodyElements.get(i);

			if (bodyElem instanceof XWPFParagraph) {
				XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
				String paragraphText = paragraph.getText();
				List<TagInfo> tags = new ArrayList<>();

				if (!DocxUtils.isNullEmpty(paragraphText)) {
					tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);

					for (TagInfo tag : tags) {
						String tagText = DocxUtils.addTagBracket(tag.getTagText());

						if (i == startCollectionIndex) {
							if (DocxUtils.isNullEmpty(startCollectionTag)) {
								if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
										|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
									startCollectionTag = DocxUtils.getTagName(tag,
											DocxConstants.TAG_PREFIX_COLLECTION_START);
								}
							}
						}

						else if (COLLECTION_END_PATTERN_1.matcher(tagText).matches()
								|| COLLECTION_END_PATTERN_2.matcher(tagText).matches()) {
							endCollectionTag = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_END);

							if (endCollectionTag.equals(startCollectionTag)) {
								collectionDO.setEndCollectionIndex(i);
								
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

				for (XWPFTableRow row : table.getRows()) {
					for (XWPFTableCell cell : row.getTableCells()) {
						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							String paragraphText = paragraph.getText();
							List<TagInfo> tags = new ArrayList<>();
							if (!DocxUtils.isNullEmpty(paragraphText)) {
								tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);
							}
						}
					}
				}
			}
		}

		return collectionDO;
	}
	

}
