package com.template.engine.tag;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.google.common.io.ByteSource;
import com.template.engine.model.CollectionDO;
import com.template.engine.model.TagInfo;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class TestProcessor {

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

	public TestProcessor(byte[] templateContent, Map<String, Object> resolutionAttributesMap) {
		super();
		this.templateContent = templateContent;
		this.resolutionAttributesMap = resolutionAttributesMap;
	}

	public TestProcessor() {
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
				//returns next element
				bodyElem = removeTagsByElement(bodyElem);
				//bodyElem = DocxUtils.getNextSibling(bodyElem);
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
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);
				for (TagInfo tag : tags) {
					process(paragraph, tag, resolutionAttributesMap, collectionDO);
				}
			}
		} else if (bodyElem instanceof XWPFTable) {

			XWPFTable table = (XWPFTable) bodyElem;

			for (XWPFTableRow row : table.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph paragraph : cell.getParagraphs()) {
						String paragraphText = paragraph.getText();
						if (!DocxUtils.isNullEmpty(paragraphText)) {
							tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);
							for (TagInfo tag : tags) {
								process(paragraph, tag, resolutionAttributesMap, collectionDO);
							}
						}
					}
				}
			}
		}
	}

	public void process(XWPFParagraph paragraph, TagInfo tag, Map<String, Object> resolutionAttributesMap,
			CollectionDO collectionDO) throws Exception {

		String tagText = DocxUtils.addTagBracket(tag.getTagText());

		if (paragraph.getText().contains(tagText)) {

			if (HEADER_PATTERN.matcher(tagText).matches() || HEADER_PATTERN_2.matcher(tagText).matches()) {
				// process header tag
				System.out.println("The header tag is: " + tagText);
				HeaderFooterTagProcessor headerFooterTag = new HeaderFooterTagProcessor();

				// before getting value, must check whether the type of the value from map is
				// object or non-object
				// returns a value after being processed
				String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_HEADER);
				String tagValue = headerFooterTag.processValue(tag, resolutionAttributesMap, tagName);

				headerFooterTag.fillHeaderFooterTag(paragraph, tagText, tagValue);
			}

			else if (FOOTER_PATTERN.matcher(tagText).matches() || FOOTER_PATTERN_2.matcher(tagText).matches()) {
				// process footer tag
				System.out.println("The header tag is: " + tagText);
				HeaderFooterTagProcessor headerFooterTag = new HeaderFooterTagProcessor();

				// before getting value, must check whether the type of the value from map is
				// object or non-object
				// returns a value after being processed
				String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_FOOTER);
				String tagValue = headerFooterTag.processValue(tag, resolutionAttributesMap, tagName);

				headerFooterTag.fillHeaderFooterTag(paragraph, tagText, tagValue);
			}

			else if (FIELD_PATTERN_1.matcher(tagText).matches() || FIELD_PATTERN_2.matcher(tagText).matches()) {
				// process field tag
				System.out.println("The field tag is: " + tagText);
				FieldTagProcessor fieldTag = new FieldTagProcessor();

				// before getting value, must check whether the type of the value from map is
				// object or non-object
				// returns a value after being processed
				String tagValue = fieldTag.processValue(tag, resolutionAttributesMap);

				fieldTag.fillFieldTag(paragraph, tagText, tagValue);
			}

			else if (IMAGE_PATTERN.matcher(tagText).matches()) {
				// process image tag
				System.out.println("The image tag is: " + tagText);
				ImageTagProcessor imageTag = new ImageTagProcessor();

				Object value = resolutionAttributesMap.get(DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_IMAGE));

				imageTag.fillImage(paragraph, tagText, value);
			}

			else if (COLLECTION_START_PATTERN_1.matcher(tagText).matches() || COLLECTION_START_PATTERN_2.matcher(tagText).matches()) {
				// process collection tag
				System.out.println("The collection tag is: " + tagText);
				CollectionTagProcessor collectionTag = new CollectionTagProcessor();
				String objectKey = null;

				// example value// -> users:user / user.phones:phone
				String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_START);
				objectKey = collectionTag.getFirstParameter(tagName); // user.phones:phone -> user.phones or listOfUser:user-> listOfUser
				//for nested collection //TODO
				if (resolutionAttributesMap.containsKey(objectKey)) {
					if (collectionDO.isHasNestedCollection()) {
						IBodyElement nextElem = DocxUtils.getNextSibling(paragraph);
						collectionTag.newProcess(nextElem, collectionDO, resolutionAttributesMap);
					} else {
						collectionDO.setTagName(tagName);
						collectionDO.setObjectFirstParameter(objectKey);
						collectionDO.setResolutionAttributesMap(resolutionAttributesMap);
						collectionDO.setTag(tag);
						collectionDO.setStartCollectionIndex(DocxUtils.getElementIndex(paragraph));
						collectionDO = collectionTag.getCollection(collectionDO);
						collectionDO = collectionTag.getEndCollectionIndex(paragraph, collectionDO.getStartCollectionIndex(), collectionDO, null);										
		
						IBodyElement nextElem = DocxUtils.getNextSibling(paragraph);
						collectionTag.newProcess(nextElem, collectionDO, resolutionAttributesMap);
					}		
				}
			}
		}
	}

	public IBodyElement removeTagsByElement(IBodyElement bodyElem) throws Exception {

		List<TagInfo> tags = new ArrayList<>();

		int elementIndex = DocxUtils.getElementIndex(bodyElem);

		// before remove get next element
		IBodyElement nextElement = DocxUtils.getNextSibling(bodyElem);

		// remove if
		if (bodyElem instanceof XWPFParagraph) {
			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			String paragraphText = paragraph.getText();

			if (!DocxUtils.isNullEmpty(paragraphText)) {
				tags = DocxUtils.getTagsByElement(paragraphText, 0, tags);

				for (TagInfo tag : tags) {
					String tagText = DocxUtils.addTagBracket(tag.getTagText());

					if (paragraph.getText().contains(tagText)) {
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
		}

		return nextElement;
	}

}
