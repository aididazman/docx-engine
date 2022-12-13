package com.template.engine;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.collections.IteratorUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.google.common.io.ByteSource;
import com.template.engine.model.CollectionDO;
import com.template.engine.model.ParentTableDO;
import com.template.engine.model.TagInfo;
import com.template.engine.tag.CollectionTagProcessor;
import com.template.engine.tag.FieldTagProcessor;
import com.template.engine.tag.HeaderFooterTagProcessor;
import com.template.engine.tag.ImageTagProcessor;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class DocxEngine {

	private static final Pattern FIELD_PATTERN_1 = Pattern.compile("\\$\\{field:[a-zA-Z]+\\}");
	private static final Pattern FIELD_PATTERN_2 = Pattern.compile("\\$\\{field:[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_1 = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_2 = Pattern.compile("\\$\\{collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_1 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_2 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern IMAGE_PATTERN = Pattern.compile("\\$\\{image:[a-zA-Z]+\\}");
	private static final Pattern HEADER_PATTERN = Pattern.compile("\\$\\{header:[a-zA-Z]+\\}");
	private static final Pattern HEADER_PATTERN_2 = Pattern.compile("\\$\\{header:[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern FOOTER_PATTERN = Pattern.compile("\\$\\{footer:[a-zA-Z]+\\}");
	private static final Pattern FOOTER_PATTERN_2 = Pattern.compile("\\$\\{footer:[a-zA-Z]+\\.[a-zA-Z]+\\}");

	private byte[] templateContent;
	private Map<String, Object> resolutionAttributesMap;

	public DocxEngine(byte[] templateContent, Map<String, Object> resolutionAttributesMap) {
		super();
		this.templateContent = templateContent;
		this.resolutionAttributesMap = resolutionAttributesMap;
	}

	public DocxEngine() {
		// TODO Auto-generated constructor stub
	}

	public byte[] generateDocument() throws Exception {
		if (templateContent == null)
			throw new Exception("Template content is null.");

		InputStream inputStream = ByteSource.wrap(templateContent).openStream();
		XWPFDocument document = new XWPFDocument(inputStream);

		CollectionDO collectionDO = new CollectionDO();

		for (XWPFHeader header : document.getHeaderList()) {
			for (IBodyElement headerElem : header.getBodyElements()) {
				processTagType(headerElem, resolutionAttributesMap, collectionDO);
			}
		}

		for (XWPFFooter footer : document.getFooterList()) {
			for (IBodyElement footerElem : footer.getBodyElements()) {
				processTagType(footerElem, resolutionAttributesMap, collectionDO);
			}
		}

		if (!DocxUtils.isNullEmpty(document.getBodyElements())) {
			IBodyElement bodyElem = document.getBodyElements().get(0);
			while (bodyElem != null) {
				processTagType(bodyElem, resolutionAttributesMap, collectionDO);
				// returns next element after removing in-replaced tags
				bodyElem = removeTagsByElement(bodyElem);
				//bodyElem = DocxUtils.getNextElement(bodyElem);
			}

		}

		ByteArrayOutputStream out = new ByteArrayOutputStream();
		document.write(out);
		out.close();
		document.close();
		inputStream.close();

		return out.toByteArray();
	}

	public void processTagType(IBodyElement bodyElem, Map<String, Object> resolutionAttributesMap,
			CollectionDO collectionDO) throws Exception {

		List<TagInfo> tags = new ArrayList<>();

		if (bodyElem instanceof XWPFParagraph) {

			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			String paragraphText = paragraph.getText();
			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, bodyElem);
				for (TagInfo tag : tags) {
					process(paragraph, tag, resolutionAttributesMap, collectionDO);
				}
			}
		} else if (bodyElem instanceof XWPFTable) {

			XWPFTable table = (XWPFTable) bodyElem;
			
			for (int row = 0; row < table.getRows().size(); row++) {
				XWPFTableRow tableRow = table.getRows().get(row);
				
				for (int cell = 0; cell < tableRow.getTableCells().size(); cell++) {
					XWPFTableCell tableCell = tableRow.getTableCells().get(cell);
					
					IBodyElement cellBodyElement = tableCell.getBodyElements().get(0);
					while (cellBodyElement != null) {					
						ParentTableDO parentTableDO = new ParentTableDO();
						parentTableDO.setTable(table);
						parentTableDO.setRowIndex(row);
						parentTableDO.setCellIndex(cell);
						
						collectionDO.setElementInTable(true);
						collectionDO.setParentTableDO(parentTableDO);
						
						processTagType(cellBodyElement, resolutionAttributesMap, collectionDO);
						cellBodyElement = removeTagsByElementInTable(tableCell, cellBodyElement);

					}
				}
			}
		}
	}
	

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void process(XWPFParagraph paragraph, TagInfo tag, Map<String, Object> resolutionAttributesMap,
			CollectionDO collectionDO) throws Exception {

		String tagText = DocxUtils.addTagBracket(tag.getTagText());

		// process header tag
		if (HEADER_PATTERN.matcher(tagText).matches() || HEADER_PATTERN_2.matcher(tagText).matches()) {

			System.out.println("The header tag is: " + tagText);
			HeaderFooterTagProcessor headerFooterTag = new HeaderFooterTagProcessor();

			// before getting value, must check whether the type of the value from map is
			// object or non-object
			// returns a value after being processed
			String tagName = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_HEADER);
			String tagValue = headerFooterTag.processValue(tagName, resolutionAttributesMap);

			headerFooterTag.fillHeaderFooterTag(paragraph, tagText, tagValue);
		}

		// process footer tag
		else if (FOOTER_PATTERN.matcher(tagText).matches() || FOOTER_PATTERN_2.matcher(tagText).matches()) {

			System.out.println("The footer tag is: " + tagText);
			HeaderFooterTagProcessor headerFooterTag = new HeaderFooterTagProcessor();

			// before getting value, must check whether the type of the value from map is
			// object or non-object
			// returns a value after being processed
			String tagName = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_FOOTER);
			String tagValue = headerFooterTag.processValue(tagName, resolutionAttributesMap);

			headerFooterTag.fillHeaderFooterTag(paragraph, tagText, tagValue);
		}

		// process field tag
		else if (FIELD_PATTERN_1.matcher(tagText).matches() || FIELD_PATTERN_2.matcher(tagText).matches()) {

			System.out.println("The field tag is: " + tagText);
			FieldTagProcessor fieldTag = new FieldTagProcessor();

			// before getting value, must check whether the type of the value from map is
			// object or non-object
			// returns a value after being processed
			String tagName = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_FIELD);
			String tagValue = fieldTag.processValue(tagName, resolutionAttributesMap);

			fieldTag.fillFieldTag(paragraph, tagText, tagValue);
		}

		// process image tag
		else if (IMAGE_PATTERN.matcher(tagText).matches()) {

			System.out.println("The image tag is: " + tagText);
			ImageTagProcessor imageTag = new ImageTagProcessor();

			Object value = resolutionAttributesMap.get(DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_IMAGE));

			imageTag.fillImage(paragraph, tagText, value);
		}

		// process collection tag
		else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
				|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {

			System.out.println("The collection tag is: " + tagText);
			CollectionTagProcessor collectionTag = new CollectionTagProcessor();
			String mapKey = null;

			// example value// -> users:user / user.phones:phone
			String tagName = DocxUtils.getTagName(tag.getTagText(), DocxConstants.TAG_PREFIX_COLLECTION_START);
			mapKey = DocxUtils.getMapKey(tagName); // user.phones:phone -> user.phones or listOfUser:user->
														// listOfUser
			// for nested collection
			if (resolutionAttributesMap.containsKey(mapKey)) {
				
				if (collectionDO.isNestedCollection()) {
					IBodyElement nextElement = DocxUtils.getNextElement(paragraph);
					collectionTag.newProcess(nextElement, collectionDO, resolutionAttributesMap);	
				} 
				
				else {
					collectionDO.setTagName(tagName);
					collectionDO.setMapKey(mapKey);
					collectionDO.setResolutionAttributesMap(resolutionAttributesMap);
					collectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(paragraph));
					collectionDO = collectionTag.getEndCollection(paragraph,
							collectionDO.getStartCollectionIndex(), collectionDO, null);

					if (resolutionAttributesMap.get(mapKey) instanceof ArrayList) {
						ListIterator<Object> iterator = ((ArrayList) resolutionAttributesMap.get(mapKey))
								.listIterator();
						List<Object> collectionValues = IteratorUtils.toList(iterator);
						collectionDO.setCollectionValues(collectionValues);
					}

					else if (resolutionAttributesMap.get(mapKey) instanceof Object) {
						collectionDO = collectionTag.getCollection(collectionDO);
					}

					IBodyElement nextElement = DocxUtils.getNextElement(paragraph);
					collectionTag.newProcess(nextElement, collectionDO, resolutionAttributesMap);
				}
			}
		}
	}

	private IBodyElement removeTagsByElement(IBodyElement bodyElem) throws Exception {

		List<TagInfo> tags = new ArrayList<>();

		int elementIndex = DocxUtils.getElementIndex(bodyElem);

		// before remove get next element
		IBodyElement nextElement = DocxUtils.getNextElement(bodyElem);

		// remove if
		if (bodyElem instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			String paragraphText = paragraph.getText();

			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, bodyElem);

				for (TagInfo tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()) {
						bodyElem.getBody().getXWPFDocument().removeBodyElement(elementIndex);
					} else if (COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
						bodyElem.getBody().getXWPFDocument().removeBodyElement(elementIndex);
					} else if (COLLECTION_END_PATTERN_1.matcher(tagText).matches()) {
						bodyElem.getBody().getXWPFDocument().removeBodyElement(elementIndex);
					} else if (COLLECTION_END_PATTERN_2.matcher(tagText).matches()) {
						bodyElem.getBody().getXWPFDocument().removeBodyElement(elementIndex);
					} else if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
						bodyElem.getBody().getXWPFDocument().removeBodyElement(elementIndex);
					}
				}
			}
		}

		return nextElement;
	}
	
	private IBodyElement removeTagsByElementInTable(XWPFTableCell tableCell, IBodyElement cellBodyElement) throws Exception {

		List<TagInfo> tags = new ArrayList<>();

		// before remove get next element
		IBodyElement nextElement = DocxUtils.getNextElement(cellBodyElement);

		// remove if
		if (cellBodyElement instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) cellBodyElement;
			String paragraphText = paragraph.getText();
			int elementIndex = DocxUtils.getParagraphIndex(tableCell.getParagraphs(), paragraph);
			boolean isLastParagraph = paragraph == tableCell.getParagraphs().get(tableCell.getParagraphs().size() - 1);

			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, cellBodyElement);

				for (TagInfo tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()) {
						if (isLastParagraph) {
							for(XWPFRun run : paragraph.getRuns()) {
								run.setText(DocxConstants.EMPTY_STRING, 0);
							}
						} else
							tableCell.removeParagraph(elementIndex);
						
					} else if (COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
						if (isLastParagraph) {
							for(XWPFRun run : paragraph.getRuns()) {
								run.setText(DocxConstants.EMPTY_STRING, 0);
							}
						} else
							tableCell.removeParagraph(elementIndex);
						
					} else if (COLLECTION_END_PATTERN_1.matcher(tagText).matches()) {
						if (isLastParagraph) {
							for(XWPFRun run : paragraph.getRuns()) {
								run.setText(DocxConstants.EMPTY_STRING, 0);
							}
						} else
							tableCell.removeParagraph(elementIndex);
						
					} else if (COLLECTION_END_PATTERN_2.matcher(tagText).matches()) {
						if (isLastParagraph) {
							for(XWPFRun run : paragraph.getRuns()) {
								run.setText(DocxConstants.EMPTY_STRING, 0);
							}
						} else
							tableCell.removeParagraph(elementIndex);
						
					} else if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
						if (isLastParagraph) {
							for(XWPFRun run : paragraph.getRuns()) {
								run.setText(DocxConstants.EMPTY_STRING, 0);
							}
						} else
							tableCell.removeParagraph(elementIndex);
					}
				}
			}
		}

		return nextElement;
	}

}
