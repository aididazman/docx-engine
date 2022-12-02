package com.template.engine.tag;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.IBody;
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
import com.template.engine.model.TagInfo;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class AllTagProcessor {

	private byte[] templateContent;
	private Map<String, Object> resolutionAttributesMap;

	private int startIndex = 0;
	private int endIndex = 0;
	private Collection tagValues = null;
	private List<TagInfo> tagsInBetween = new ArrayList<>();

	private static final Pattern FIELD_PATTERN = Pattern.compile("\\$\\{field:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_START_PATTERN = Pattern.compile("\\$\\{collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_END_PATTERN = Pattern.compile("\\$\\{/collection:[a-zA-Z]+:[a-zA-Z]+\\}");
	private static final Pattern COLLECTION_OBJECT_PATTERN = Pattern.compile("\\$\\{[a-zA-Z]+\\.[a-zA-Z]+\\}");
	private static final Pattern IMAGE_PATTERN = Pattern.compile("\\$\\{image:[a-zA-Z]+\\}");
	private static final Pattern HEADER_PATTERN = Pattern.compile("\\$\\{header:[a-zA-Z]+\\}");
	private static final Pattern FOOTER_PATTERN = Pattern.compile("\\$\\{footer:[a-zA-Z]+\\}");

	public AllTagProcessor(byte[] templateContent, Map<String, Object> resolutionAttributesMap) {
		super();
		this.templateContent = templateContent;
		this.resolutionAttributesMap = resolutionAttributesMap;
	}

	public byte[] generateDocument() throws Exception {
		if (templateContent == null)
			throw new Exception("Template content is null.");

		InputStream inputStream = ByteSource.wrap(templateContent).openStream();
		XWPFDocument document = new XWPFDocument(inputStream);

		List<TagInfo> tags = new ArrayList<>();
		List<TagInfo> headerTags = new ArrayList<>();
		List<TagInfo> footerTags = new ArrayList<>();
		CollectionDO collectionDO = new CollectionDO();

		int counterHeader = 0;
		for (XWPFHeader header : document.getHeaderList()) {
			for (IBodyElement headerElem : header.getBodyElements()) {
				counterHeader = processTagType(headerElem, headerTags, resolutionAttributesMap, collectionDO, counterHeader);
			}
		}

		int counterFooter = 0;
		for (XWPFFooter footer : document.getFooterList()) {
			for (IBodyElement footerElem : footer.getBodyElements()) {
				counterFooter = processTagType(footerElem, footerTags, resolutionAttributesMap, collectionDO, counterFooter);
			}
		}

		//test document.removeBodyElement();
//		List<IBodyElement> bodyElements = document.getBodyElements();
//		IBodyElement bodyEleme = null;
//		bodyEleme.getBody();
//		XWPFParagraph paragraph = null;
//		paragraph.getDocument();		
		
		int counter = 0;
		// replace tags in the document body
		for (IBodyElement bodyElem : document.getBodyElements()) {
			counter = processTagType(bodyElem, tags, resolutionAttributesMap, collectionDO, counter);
		}

		ByteArrayOutputStream out = new ByteArrayOutputStream();
		document.write(out);
		out.close();
		document.close();
		inputStream.close();

		return out.toByteArray();
	}

//	public List<TagInfo> getAllTags(IBody document, List<TagInfo> tags) throws Exception {
//
//		String text = getAllText(document.getBodyElements());
//
//		if (!DocxUtils.isNullEmpty(text)) {
//			tags = getTagsByElement(text, 0, tags);
//		}
//
//		return tags;
//	}
//
//	private String getAllText(List<IBodyElement> bodyElements) {
//
//		StringBuilder text = new StringBuilder();
//
//		for (int i = 0; i < bodyElements.size(); i++) {
//			getTextByElement(text, bodyElements.get(i));
//		}
//
//		return text.toString();
//	}

//	private void getTextByElement(StringBuilder text, IBodyElement bodyElem) {
//
//		if (bodyElem instanceof XWPFParagraph) {
//			text.append(((XWPFParagraph) bodyElem).getText());
//
//		} else if (bodyElem instanceof XWPFTable) {
//			XWPFTable table = (XWPFTable) bodyElem;
//
//			for (XWPFTableRow row : table.getRows()) {
//
//				for (XWPFTableCell cell : row.getTableCells()) {
//					List<IBodyElement> localBodyElements = cell.getBodyElements();
//
//					for (int i = 0; i < localBodyElements.size(); i++) {
//						getTextByElement(text, localBodyElements.get(i));
//					}
//				}
//			}
//		}
//	}

	private List<TagInfo> getTagsByElement(String elementText, int tagStartOffset, List<TagInfo> tags)
			throws Exception {

		tagStartOffset = elementText.indexOf(DocxConstants.DEFAULT_TAG_START, tagStartOffset);
		if (tagStartOffset >= 0) {
			int tagEndOffset = elementText.indexOf(DocxConstants.DEFAULT_TAG_END, tagStartOffset);

			if (tagEndOffset < 0) {
				throw new Exception("No closing tag found for line " + elementText);
			}

			String tagText = elementText.substring(tagStartOffset + DocxConstants.DEFAULT_TAG_START.length(),
					tagEndOffset);
			boolean hasClosingSlash = tagText.startsWith("/");

			TagInfo tagInfo = new TagInfo(tagText, tagEndOffset, hasClosingSlash);
			tags.add(tagInfo);

			// recursive to proceed to get other tags in the same text
			getTagsByElement(elementText, tagEndOffset, tags);
		}

		return tags;
	}

	private int processTagType(IBodyElement bodyElem, List<TagInfo> tags, Map<String, Object> resolutionAttributesMap,
			CollectionDO collectionDO, int counter) throws Exception {

		if (bodyElem instanceof XWPFParagraph) {
			counter = processEachTag(((XWPFParagraph) bodyElem), tags, resolutionAttributesMap, collectionDO, counter);
		} else if (bodyElem instanceof XWPFTable) {
			counter = processEachTag(((XWPFTable) bodyElem), tags, resolutionAttributesMap, collectionDO, counter);
		}
		return counter;
	}

	// must return
	private int processEachTag(XWPFParagraph paragraph, List<TagInfo> tags,
			Map<String, Object> resolutionAttributesMap, CollectionDO collectionDO, int counter) throws Exception {

		String paragraphText = paragraph.getText();
		if (!DocxUtils.isNullEmpty(paragraphText)) {
			tags = getTagsByElement(paragraphText, 0, tags);
			for (TagInfo tag : tags) {
				counter = process(paragraph, tag, resolutionAttributesMap, collectionDO, counter);
			}
		}
		return counter;
	}

	private int processEachTag(XWPFTable table, List<TagInfo> tags, Map<String, Object> resolutionAttributesMap,
			CollectionDO collectionDO, int counter) throws Exception {

		for (XWPFTableRow row : table.getRows()) {
			for (XWPFTableCell cell : row.getTableCells()) {
				for (XWPFParagraph paragraph : cell.getParagraphs()) {
					String paragraphText = paragraph.getText();
					if (!DocxUtils.isNullEmpty(paragraphText)) {
						tags = getTagsByElement(paragraphText, 0, tags);
						for (TagInfo tag : tags) {
							counter = process(paragraph, tag, resolutionAttributesMap, collectionDO, counter);
						}
					}
				}
			}
		}
		return counter;
	}

	private int process(XWPFParagraph paragraph, TagInfo tag, Map<String, Object> resolutionAttributesMap,
			CollectionDO collectionDO, int counter) throws Exception {

		String tagText = DocxUtils.addTagBracket(tag.getTagText());

		if (paragraph.getText().contains(tagText)) {

			if (HEADER_PATTERN.matcher(tagText).matches()) {
				// process header tag
				System.out.println("The header tag is: " + tagText);
				HeaderFooterTagProcessor headerFooterTag = new HeaderFooterTagProcessor();

				Object value = resolutionAttributesMap.get(DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_HEADER));
				String tagValue = headerFooterTag.getValue(tag, value, DocxConstants.TAG_PREFIX_HEADER);

				headerFooterTag.fillHeaderFooterTag(paragraph, tagText, tagValue);
			}

			else if (FOOTER_PATTERN.matcher(tagText).matches()) {
				// process footer tag
				System.out.println("The header tag is: " + tagText);
				HeaderFooterTagProcessor headerFooterTag = new HeaderFooterTagProcessor();

				Object value = resolutionAttributesMap.get(DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_FOOTER));
				String tagValue = headerFooterTag.getValue(tag, value, DocxConstants.TAG_PREFIX_FOOTER);

				headerFooterTag.fillHeaderFooterTag(paragraph, tagText, tagValue);
			}

			else if (FIELD_PATTERN.matcher(tagText).matches()) {
				// process field tag
				System.out.println("The field tag is: " + tagText);
				FieldTagProcessor fieldTag = new FieldTagProcessor();

				String tagValue = fieldTag.checkMapValueType(tag, resolutionAttributesMap);

				fieldTag.fillFieldTag(paragraph, tagText, tagValue);
			}

			else if (IMAGE_PATTERN.matcher(tagText).matches()) {
				// process image tag
				System.out.println("The image tag is: " + tagText);
				ImageTagProcessor imageTag = new ImageTagProcessor();

				Object value = resolutionAttributesMap.get(DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_IMAGE));

				imageTag.fillImage(paragraph, tagText, value);
			}

			else if (COLLECTION_START_PATTERN.matcher(tagText).matches()) {
				// process collection tag
				System.out.println("The collection tag is: " + tagText);
				CollectionTagProcessor collectionTag = new CollectionTagProcessor();

				// 1st, we take 1st parameter from the collection tag, example->
				// collection:users:user, 1st parameter is users
				// 2nd we check whether users inside the resolution map

				// example value// -> users:user
				String tagName = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_COLLECTION_START);

				collectionDO.setTagName(tagName);
				collectionDO.setResolutionAttributesMap(resolutionAttributesMap);
				collectionDO.setTag(tag);
				collectionDO = collectionTag.getCollection(collectionDO);

				// replace ${collection:users:user} with empty text
				// DocxUtils.replaceTextSegment(paragraph, tagText, DocxConstants.EMPTY_STRING);
			}

			else if (COLLECTION_OBJECT_PATTERN.matcher(tagText).matches()) {

				CollectionTagProcessor collectionTag = new CollectionTagProcessor();
				// first check whether tag object name matches with collection value
				if (!DocxUtils.isNullEmpty(collectionDO.getCollectionValues())) {
					for (Object collectionValue : collectionDO.getCollectionValues()) {

						String tagObjectName = collectionTag.getTagObjectName(tag); // example value from user.name ->
																					// user
						if (collectionValue.getClass().getSimpleName().equalsIgnoreCase(tagObjectName)) {
							// then we take tag object field from tag.getText()
							String tagObjectField = collectionTag.getTagObjectField(tag); // example value from
																							// user.name -> name
							// get the value of the field based on the tag object field
							String value = collectionTag.getFieldValue(collectionValue, tagObjectField);
							// replace tag text with
							collectionTag.insertNewParagraph(paragraph, tagText, value, tag, 0);

						}
					}
				}
			}

			/*
			 * CollectionTagProcessor collectionTag = new CollectionTagProcessor();
			 * 
			 * if (COLLECTION_START_PATTERN.matcher(tagText).matches()) { // process
			 * collection tag System.out.println("The collection tag is: " + tagText);
			 * startIndex = i;// to do-> map to store start and end index
			 * 
			 * // no need loop for (Entry<String, Object> entry :
			 * resolutionAttributesMap.entrySet()) {
			 * 
			 * if (entry.getKey() .contains(DocxUtils.getTagName(tag,
			 * DocxConstants.TAG_PREFIX_COLLECTION_START))) {
			 * 
			 * tagValues = collectionTag.getValue(tag, entry.getValue(), tagValues); }
			 * 
			 * }
			 * 
			 * }
			 * 
			 * if (COLLECTION_END_PATTERN.matcher(tagText).matches()) { // process
			 * collection tag System.out.println("The end collection tag is: " + tagText);
			 * endIndex = i; // get tags between collection start and end tag tagsInBetween
			 * = collectionTag.getTagsInBetween(startIndex, endIndex, tags, tagsInBetween);
			 * collectionTag.process(bodyElements, tagsInBetween, tagValues);
			 * 
			 * }
			 */

		}

		return counter += 1;
	}

}
