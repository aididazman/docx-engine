package com.template.engine.tag;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;

import com.template.engine.model.CollectionDO;
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
			;
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

	private String getFieldValue(Object mapValue, String objectField) {

		Object value = null;
		try {
			value = PropertyUtils.getSimpleProperty(mapValue, objectField);
		} catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
			throw new RuntimeException("Cannot get tag " + objectField + " value from the context");
		}
		String tagValue = value == null ? null : value.toString();
		if (tagValue == null) {
			tagValue = DocxConstants.EMPTY_STRING;
		}

		return tagValue;
	}

	private CollectionDO getCollectionValues(CollectionDO collectionDO) {

		// check type pattern
		if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(collectionDO.getTagName()).matches()) {
			// collection tag -> collection:user.phones:phone
			// returns example value user.phones:phone -> user.phones
			String collectionObjectKey = collectionDO.getObjectFirstParameter();
			// returns phones from user.phones
			String collectionObjectName = getClassField(collectionObjectKey);
			// get collection values from map object value where key is user.phones and set
			// it in collectionDO object
			collectionDO.setCollectionValues(getCollectionValues(collectionObjectName,
					collectionDO.getResolutionAttributesMap().get(collectionObjectKey)));
		}

		else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(collectionDO.getTagName()).matches()) {
			// take object name from 1st parameter
			// example value extracted from listOfUser:users = listOfUser
			String collectionObjectName = collectionDO.getObjectFirstParameter();
			// returns users from listOfUser:users
			String objectField = getSecondParameter(collectionDO.getTagName());
			collectionDO.setCollectionValues(getCollectionValues(objectField,
					collectionDO.getResolutionAttributesMap().get(collectionObjectName)));
		}

		return collectionDO;
	}

	private void processCollectionInParagraph(XWPFParagraph elementAfterStartCollection, CollectionDO collectionDO,
			Map<String, Object> resolutionAttributesMap) throws Exception {

		List<TagInfo> tags = new ArrayList<>();
		String nonParentNestedCollectionName = null;

		XWPFParagraph paragraph = (XWPFParagraph) elementAfterStartCollection;
				
		tags = getTagsFromCollection(elementAfterStartCollection.getBody(), tags, collectionDO.getStartCollectionIndex() + 1, collectionDO.getEndCollectionIndex());

		if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {

			for (Object collectionValue : collectionDO.getCollectionValues()) {

				for (TagInfo tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
						// example value from user.name -> user
						String className = getClassName(tag.getTagText());
						if (collectionValue.getClass().getSimpleName().equalsIgnoreCase(className)) {
							// then we take tag object field from tag.getText()
							// example value from user.name -> name
							String objectField = getClassField(tag.getTagText());
							// get the value of the field based on the tag object field
							String value = getFieldValue(collectionValue, objectField);
							boolean isLastTag = tag == tags.get(tags.size() - 1);
							// replace tag text with
							insertNewParagraph(paragraph, value, collectionDO.getEndCollectionElement(), isLastTag);

						}
					}

					else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
						String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_START);
						TestProcessor testProcessor = new TestProcessor();

						if (COLLECTION_OBJECT_FIELD_PATTERN_1.matcher(tagName).matches()) {
							// returns example value user.phones:phone -> user.phones
							String nestedCollectionName = getFirstParameter(tagName);
							// returns user from user.phones
							String nestedCollectionClassName = getClassName(nestedCollectionName);
							// returns example value user.phones -> phones -> change key to user.phones
							// String tagObjectName = getObjectNameTypeTwo(nestedCollectionObjectName);

							// belongs to parent collection values
							if (collectionValue.getClass().getSimpleName().equalsIgnoreCase(nestedCollectionClassName)) {
								// new map
								Map<String, Object> newValues = new HashMap<String, Object>();
								newValues.put(nestedCollectionName, collectionValue);
								
								IBodyElement nestedCollectionElement = elementAfterStartCollection.getBody()
										.getBodyElements().get(tag.getTagElementIndex());
								
								CollectionDO nestedCollectionDO = new CollectionDO();
								
								boolean isLastCollectionValue = collectionValue == collectionDO.getCollectionValues().get(collectionDO.getCollectionValues().size() - 1);
								if (isLastCollectionValue) {
									nestedCollectionDO.setLastCollectionValue(true);
								} else
									nestedCollectionDO.setLastCollectionValue(false);
								
								nestedCollectionDO.setTagName(tagName);
								nestedCollectionDO.setObjectFirstParameter(nestedCollectionName);
								nestedCollectionDO.setResolutionAttributesMap(newValues);
								nestedCollectionDO.setTag(tag);
								nestedCollectionDO.setStartCollectionIndex(tag.getTagElementIndex());
								nestedCollectionDO.setHasNestedCollection(true);
								nestedCollectionDO = getEndCollectionIndex(paragraph, tag.getTagElementIndex(), nestedCollectionDO, collectionDO.getEndCollectionElement());	
								nestedCollectionDO = getCollection(nestedCollectionDO);
																							
								if (nestedCollectionElement instanceof XWPFParagraph) {
									XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
									testProcessor.process(nestedParagraph, tag, newValues, nestedCollectionDO);
								} 
							}
						}

						else if (COLLECTION_OBJECT_FIELD_PATTERN_2.matcher(tagName).matches()) {
							nonParentNestedCollectionName = getFirstParameter(tagName); // employees:name -> employees

							//nestedCollectionDO.setEndCollectionElement(collectionDO.getEndCollectionElement());
							
							if (resolutionAttributesMap.containsKey(nonParentNestedCollectionName)) {
								Object newCollectionValue = resolutionAttributesMap.get(nonParentNestedCollectionName);
								CollectionDO nestedCollectionDO = new CollectionDO();
								
								Map<String, Object> newValues = new HashMap<String, Object>();
								newValues.put(nonParentNestedCollectionName, newCollectionValue);
								
								nestedCollectionDO.setTagName(tagName);
								nestedCollectionDO.setObjectFirstParameter(nonParentNestedCollectionName);
								nestedCollectionDO.setResolutionAttributesMap(newValues); //set resolution attribute
								nestedCollectionDO.setTag(tag);
								nestedCollectionDO.setStartCollectionIndex(tag.getTagElementIndex());
								nestedCollectionDO.setHasNestedCollection(true);								
								nestedCollectionDO = getEndCollectionIndex(paragraph, tag.getTagElementIndex(), nestedCollectionDO, collectionDO.getEndCollectionElement());	
								nestedCollectionDO = getCollection(nestedCollectionDO);
								
								IBodyElement nestedCollectionElement = elementAfterStartCollection.getBody()
										.getBodyElements().get(tag.getTagElementIndex());

								if (nestedCollectionElement instanceof XWPFParagraph) {
									XWPFParagraph nestedParagraph = (XWPFParagraph) nestedCollectionElement;
									testProcessor.process(nestedParagraph, tag, newValues, nestedCollectionDO);									
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
	
	private void processCollectionInTableNested(XWPFTable tableElementAfterStartCollection, CollectionDO collectionDO) throws Exception {
		
		XWPFTable parentTable = (XWPFTable) tableElementAfterStartCollection;
		
		IBodyElement endCollectionElement = collectionDO.getEndCollectionElement();
		XWPFTable newTable = insertNewTable(parentTable, endCollectionElement);
		
		if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {
			
			int newColumnSize = collectionDO.getCollectionValues().get(0).getClass().getDeclaredFields().length;
			int newRowSize = collectionDO.getCollectionValues().size();
			
			for (int column = 1; column < newColumnSize; column++) {
				newTable.getRow(0).createCell();
			}
			
			for (int row = 0; row < newRowSize; row++) {
				newTable.createRow();
			}

			for (int row = 0; row < newTable.getRows().size(); row++) {
				XWPFTableRow tableRow = newTable.getRows().get(row);
				
				for (int cell = 0; cell < tableRow.getTableCells().size(); cell++) {
					XWPFTableCell tableCell = tableRow.getTableCells().get(cell);
					// first row for each column text of parent table
					String firstRowText = parentTable.getRow(1).getCell(cell).getTextRecursively();
					List<TagInfo> tags = new ArrayList<>();
					tags = DocxUtils.getTagsByElement(firstRowText, 0, tags);
					
					for (XWPFParagraph paragraph : tableCell.getParagraphs()) {
						
						for (TagInfo tag : tags) {
							//set headers of the table
							if (row == 0) {
								XWPFRun run = paragraph.createRun();
								String text = run.getText(0);
								
								if (DocxUtils.isNullEmpty(text))
									text = DocxConstants.EMPTY_STRING;
								
								String headerText = parentTable.getRow(0).getCell(cell).getTextRecursively(); // set in value to be replaced
								text = text.replace(text, headerText);
								run.setText(headerText, 0);
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
										String className = getClassName(tag.getTagText());

										if (collectionValue.getClass().getSimpleName().equalsIgnoreCase(className)) {
											String objectField = getClassField(tag.getTagText());
											String value = getFieldValue(collectionValue, objectField);
											text = text.replace(text, value);
											run.setText(value, 0);
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
	

	private void processCollectionInTable(XWPFTable table, CollectionDO collectionDO)
			throws Exception {
		
		if (collectionDO.isHasNestedCollection()) {
			processCollectionInTableNested(table, collectionDO);
		} else {
			for (int value = 0; value < collectionDO.getCollectionValues().size(); value++) {
				table.createRow();
			}

			for (int row = 2; row < table.getRows().size(); row++) {

				XWPFTableRow rowTable = table.getRows().get(row);

				for (int cell = 0; cell < rowTable.getTableCells().size(); cell++) {
					XWPFTableCell cellTable = rowTable.getTableCells().get(cell);
					// first row for each column text
					String firstRowText = table.getRow(1).getCell(cell).getTextRecursively();
					List<TagInfo> tags = new ArrayList<>();
					tags = DocxUtils.getTagsByElement(firstRowText, 0, tags);

					for (XWPFParagraph paragraph : cellTable.getParagraphs()) {

						for (TagInfo tag : tags) {

							String tagText = DocxUtils.addTagBracket(tag.getTagText());

							if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {

								XWPFRun run = paragraph.createRun();
								String text = run.getText(0);

								if (DocxUtils.isNullEmpty(text))
									text = DocxConstants.EMPTY_STRING;

								Object collectionValue = collectionDO.getCollectionValues().get(row - 2);
								String className = getClassName(tag.getTagText());

								if (collectionValue.getClass().getSimpleName().equalsIgnoreCase(className)) {
									String objectField = getClassField(tag.getTagText());
									String value = getFieldValue(collectionValue, objectField);
									text = text.replace(text, value);
									run.setText(value, 0);
								}
							}
						}
					}
				}
			}

			table.removeRow(1);
		}
	}
	

	
	private void insertNewParagraph(XWPFParagraph paragraph, String value, IBodyElement endCollectionElement, boolean isLastTag)
			throws Exception {

		XWPFDocument document = paragraph.getDocument();

		XWPFParagraph endTagElement = (XWPFParagraph) endCollectionElement;
		XmlCursor cursor = endTagElement.getCTP().newCursor();
		XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
		XWPFRun run = newParagraph.createRun();
		run.setText(value);
		
		if (isLastTag) {
			run.addBreak(BreakType.TEXT_WRAPPING);
		}
		
	}
	
	private XWPFTable insertNewTable(XWPFTable table, IBodyElement endCollectionElement)
			throws Exception {

		XWPFDocument document = table.getBody().getXWPFDocument();

		XWPFParagraph endTagElement = (XWPFParagraph) endCollectionElement;
		XmlCursor cursor = endTagElement.getCTP().newCursor();
		XWPFTable newTable = document.insertNewTbl(cursor);
		
		return newTable;
		// run.addBreak(BreakType.TEXT_WRAPPING);
	}

	private List<TagInfo> getTagsFromCollection(IBody document, List<TagInfo> tags, int startIndex, int endIndex) throws Exception {

		List<IBodyElement> subListBodyElements = document.getBodyElements();
		for (int i = startIndex; i < endIndex; i++) {
			IBodyElement bodyElem = subListBodyElements.get(i);

			if (bodyElem instanceof XWPFParagraph) {
				XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
				String paragraphText = paragraph.getText();

				if (!DocxUtils.isNullEmpty(paragraphText)) {
					tags = getTagsInBetweenCollection(paragraphText, 0, i, tags);
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

	public CollectionDO getEndCollectionIndex(IBodyElement elementAfterStartCollection, int startIndex,
			CollectionDO collectionDO, IBodyElement parentCollection) throws Exception {

		List<IBodyElement> bodyElements = elementAfterStartCollection.getBody().getBodyElements();
		List<TagInfo> tags = new ArrayList<>();
		int endTagIndex = 0;
		String startCollectionTag = null;
		String endCollectionTag = null;

		for (int i = startIndex; i < bodyElements.size(); i++) {
			IBodyElement bodyElem = bodyElements.get(i);

			if (bodyElem instanceof XWPFParagraph) {
				XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
				String paragraphText = paragraph.getText();

				if (!DocxUtils.isNullEmpty(paragraphText)) {
					tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);

					for (TagInfo tag : tags) {
						String tagNameText = DocxUtils.addTagBracket(tag.getTagText());

						if (paragraphText.contains(tagNameText)) {

							if (i == startIndex) {
								if (DocxUtils.isNullEmpty(startCollectionTag)) {
									if (COLLECTION_START_PATTERN_1.matcher(tagNameText).matches()
											|| COLLECTION_START_PATTERN_2.matcher(tagNameText).matches()) {
										startCollectionTag = DocxUtils.getTagName(tag,
												DocxConstants.TAG_PREFIX_COLLECTION_START);
									}
								}
							}

							else if (COLLECTION_END_PATTERN_1.matcher(tagNameText).matches()
									|| COLLECTION_END_PATTERN_2.matcher(tagNameText).matches()) {
								endCollectionTag = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_END);

								if (endCollectionTag.equals(startCollectionTag)) {
									endTagIndex = i;
									collectionDO.setEndCollectionIndex(endTagIndex);
									// must set a new condition to cater nested collection end element
									if (!collectionDO.isHasNestedCollection()) {
										collectionDO.setEndCollectionElement(bodyElem);
									} else
										collectionDO.setEndCollectionElement(parentCollection);							
								}
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
