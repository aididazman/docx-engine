package com.docx.impl;

import java.io.ByteArrayInputStream;
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

import com.docx.model.CollectionDO;
import com.docx.model.DocxVO;
import com.docx.model.ParentTableDO;
import com.docx.model.TagDO;
import com.docx.tagprocessor.CollectionTagProcessor;
import com.docx.tagprocessor.FieldTagProcessor;
import com.docx.tagprocessor.ImageTagProcessor;
import com.docx.utils.DocxConstants;
import com.docx.utils.DocxUtils;

public class DocxEngine {

	/**
	 * Docx Engine
	 * <br>Changes History : 
	 * 
	 * @author Muhammad Aidid Azman
	 * @collaborator Muhammad Amin
	 * @version 1.0
	 * @since 5/1/2022 
	 */
	
	private static final Pattern FIELD_PATTERN_1 = Pattern.compile("\\$\\{field:[a-zA-Z]+\\}");
	private static final Pattern FIELD_PATTERN_2 = Pattern.compile("\\$\\{field:[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_1 = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN_2 = Pattern.compile("\\$\\{collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_1 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN_2 = Pattern.compile("\\$\\{/collection:[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern IMAGE_PATTERN = Pattern.compile("\\$\\{image:[a-zA-Z]+\\}");

	private byte[] templateContent;
	private Map<String, Object> mapValues;

	public DocxEngine(byte[] templateContent, Map<String, Object> mapValues) {
		super();
		this.templateContent = templateContent;
		this.mapValues = mapValues;
	}
	
	public byte[] generateDocument() throws Exception {
		if (templateContent == null)
			throw new Exception("Template content is null.");

		InputStream inputStream = new ByteArrayInputStream(templateContent);
		XWPFDocument document = new XWPFDocument(inputStream);

		DocxVO docxVO = new DocxVO();
		docxVO.setMapValues(mapValues);

		for (XWPFHeader header : document.getHeaderList()) {
			for (IBodyElement headerElem : header.getBodyElements()) {
				docxVO.setBodyElement(headerElem);
				docxVO = processTagType(docxVO);
			}
		}

		for (XWPFFooter footer : document.getFooterList()) {
			for (IBodyElement footerElem : footer.getBodyElements()) {
				docxVO.setBodyElement(footerElem);
				docxVO = processTagType(docxVO);
			}
		}

		if (!DocxUtils.isNullEmpty(document.getBodyElements())) {
			IBodyElement bodyElem = document.getBodyElements().get(0);
			while (bodyElem != null) {
				docxVO.setBodyElement(bodyElem);
				docxVO = processTagType(docxVO);
				bodyElem = removeTagsByElement(bodyElem, null, false);
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

	public DocxVO processTagType(DocxVO docxVO) throws Exception {

		List<TagDO> tagDOList = new ArrayList<>();

		if (docxVO.getBodyElement() instanceof XWPFParagraph) {

			XWPFParagraph paragraph = (XWPFParagraph) docxVO.getBodyElement();
			String paragraphText = paragraph.getText();
			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tagDOList = DocxUtils.getTagsByElement(paragraphText, 0, tagDOList, docxVO.getBodyElement());
				for (TagDO tagDO : tagDOList) {
					docxVO.setTagDO(tagDO);
					docxVO = process(docxVO);
				}
			}
		} else if (docxVO.getBodyElement() instanceof XWPFTable) {

			XWPFTable table = (XWPFTable) docxVO.getBodyElement();
			
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

						docxVO.setIsCollectionInTable(true);
						docxVO.setParentTableDO(parentTableDO);
						docxVO.setBodyElement(cellBodyElement);
		
						boolean isElementInTable = docxVO.isCollectionInTable();
						docxVO = processTagType(docxVO);
						cellBodyElement = removeTagsByElement(cellBodyElement, tableCell, isElementInTable);
						//cellBodyElement = DocxUtils.getNextElement(cellBodyElement);
					}
				}
			}
		}
		return docxVO;
	}
	

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public DocxVO process(DocxVO docxVO) throws Exception {

		XWPFParagraph paragraph = (XWPFParagraph) docxVO.getBodyElement();
		TagDO tagDO = docxVO.getTagDO();
		Map<String, Object> mapValues = docxVO.getMapValues();
		
		String tagText = DocxUtils.addTagBracket(tagDO.getTagText());

		// process field tag
		if (FIELD_PATTERN_1.matcher(tagText).matches() || FIELD_PATTERN_2.matcher(tagText).matches()) {

			System.out.println("The field tag is: " + tagText);
			
			FieldTagProcessor fieldTag = new FieldTagProcessor();
			String tagName = DocxUtils.getTagName(tagDO.getTagText(), DocxConstants.TAG_PREFIX_FIELD);
			String tagValue = fieldTag.processValue(tagName, mapValues);

			fieldTag.fillTag(paragraph, tagText, tagValue);
		}

		// process image tag
		else if (IMAGE_PATTERN.matcher(tagText).matches()) {

			System.out.println("The image tag is: " + tagText);
			
			ImageTagProcessor imageTag = new ImageTagProcessor();
			Object value = mapValues.get(DocxUtils.getTagName(tagDO.getTagText(), DocxConstants.TAG_PREFIX_IMAGE));
			imageTag.fillImage(paragraph, tagText, value);
		}

		// process collection tag
		else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
				|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {

			System.out.println("The collection tag is: " + tagText);
			CollectionTagProcessor collectionTag = new CollectionTagProcessor();
			String mapKey = null;

			// example value// -> users:user / user.phones:phone
			String tagName = DocxUtils.getTagName(tagDO.getTagText(), DocxConstants.TAG_PREFIX_COLLECTION_START);
			// user.phones:phone -> user.phones or listOfUser:user->listOfUser
			mapKey = DocxUtils.getMapKey(tagName); 
	
			if (mapValues.containsKey(mapKey)) {
				// to check whether the collection tag has been processed before
				List<IBodyElement> bodyElements = paragraph.getBody().getBodyElements();
				boolean isCollectionProcessed = false;
				isCollectionProcessed = isCollectionProcessed(isCollectionProcessed, bodyElements, docxVO);
				
				if(!isCollectionProcessed) {
					
					CollectionDO collectionDO = new CollectionDO();
					if (docxVO.isCollectionInTable()) {
						collectionDO.setElementInTable(true);
						collectionDO.setParentTableDO(docxVO.getParentTableDO());
					}				
					collectionDO.setTagName(tagName);
					collectionDO.setMapKey(mapKey);
					collectionDO.setResolutionAttributesMap(mapValues);
					collectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(paragraph));
					collectionDO = collectionTag.getEndCollection(paragraph,
							collectionDO.getStartCollectionIndex(), collectionDO, null);
					collectionDO.setNextElement(DocxUtils.getNextElement(paragraph));

					if (mapValues.get(mapKey) instanceof ArrayList) {
						ListIterator<Object> iterator = ((ArrayList) mapValues.get(mapKey)).listIterator();
						List<Object> collectionValues = IteratorUtils.toList(iterator);
						collectionDO.setCollectionValues(collectionValues);
					}

					else if (mapValues.get(mapKey) instanceof Object) {
						collectionDO = collectionTag.getCollection(collectionDO);
					}

					docxVO.setCollectionDO(collectionDO);
					collectionTag.newProcess(collectionDO);
				}
			}
		}
		
		return docxVO;
	}

	private boolean isCollectionProcessed(boolean isProcessed, List<IBodyElement> bodyElements, DocxVO docxVO) {

		if (docxVO.getCollectionDO() != null) {
			
			for (int i = 0; i < bodyElements.size(); i++) {
				IBodyElement bodyElem = bodyElements.get(i);
				if (bodyElem == docxVO.getCollectionDO().getEndCollectionElement())
					return true;
			}
		}
		
		return isProcessed;
	}

	private IBodyElement removeTagsByElement(IBodyElement bodyElem, XWPFTableCell tableCell, boolean isElementInCell)
			throws Exception {

		List<TagDO> tags = new ArrayList<>();
		// before remove get next element
		IBodyElement nextElement = DocxUtils.getNextElement(bodyElem);

		// remove if
		if (bodyElem instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			int elementIndex = isElementInCell ? DocxUtils.getParagraphIndex(tableCell.getParagraphs(), paragraph)
					: DocxUtils.getElementIndex(paragraph);
			String paragraphText = paragraph.getText();

			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags, bodyElem);

				for (TagDO tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (COLLECTION_START_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_START_PATTERN_2.matcher(tagText).matches()
							|| COLLECTION_END_PATTERN_1.matcher(tagText).matches()
							|| COLLECTION_END_PATTERN_2.matcher(tagText).matches()
							|| COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {
						removeTag(paragraph, tableCell, isElementInCell, elementIndex);
					}
				}
			}
		}

		return nextElement;
	}
	
	private void removeTag(XWPFParagraph paragraph, XWPFTableCell tableCell, boolean isElementInCell,
			int elementIndex) {

		if (isElementInCell) {
			boolean isLastParagraph = paragraph == tableCell.getParagraphs().get(tableCell.getParagraphs().size() - 1);
			if (isLastParagraph) {
				for (XWPFRun run : paragraph.getRuns()) {
					run.setText(DocxConstants.EMPTY_STRING, 0);
				}
			} else
				tableCell.removeParagraph(elementIndex);
		} else
			paragraph.getBody().getXWPFDocument().removeBodyElement(elementIndex);
	}

}
