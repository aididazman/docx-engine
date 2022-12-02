package com.template.engine.tag;

import java.lang.reflect.InvocationTargetException;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.template.engine.model.TagInfo;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class FieldTagProcessor {

	private static final Pattern FIELD_OBJECT_PATTERN = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+");

	private String getTagValue(String tagObjectField, Object mapValue) {

		Object value = null;
		try {
			value = PropertyUtils.getSimpleProperty(mapValue, tagObjectField);
		} catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
			throw new RuntimeException("Cannot get tag " + tagObjectField + " value from the context");
		}
		String tagValue = value == null ? null : value.toString(); // call framework stringbuilder combine to string
																	// method
		if (tagValue == null) {
			tagValue = DocxConstants.EMPTY_STRING;
		}

		return tagValue;
	}

	public void fillFieldTag(IBodyElement bodyElem, String textToBeReplaced, String textValue) {

		if (bodyElem instanceof XWPFParagraph) {

			XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
			DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
		} else if (bodyElem instanceof XWPFTable) {
			XWPFTable table = (XWPFTable) bodyElem;

			for (XWPFTableRow row : table.getRows()) {

				for (XWPFTableCell cell : row.getTableCells()) {

					for (XWPFParagraph paragraph : cell.getParagraphs()) {
						DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
					}

				}

			}

		}
	}

	public String processValue(TagInfo tag, Map<String, Object> resolutionAttributesMap) {

		String value = null;
		Object mapValue = null;
		// get tag text from tag, example field:name / field:user.name -> name /
		// user.name
		String tagText = DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_FIELD);
		// check whether tag text match type pattern
		if (FIELD_OBJECT_PATTERN.matcher(tagText).matches()) {
			String tagObjectName = getFirstParameterTypeTwo(tagText);// get user from user.name
			mapValue = resolutionAttributesMap.get(tagText);
			// check whether tag object name equals the name of the class from map value
			if (mapValue.getClass().getSimpleName().equalsIgnoreCase(tagObjectName)) {
				String tagObjectField = getSecondParameterTypeTwo(tagText); // get name from user.name
				value = getTagValue(tagObjectField, mapValue);
			}
		}

		else {
			mapValue = resolutionAttributesMap.get(tagText);
			value = DocxUtils.combineToString(mapValue);

		}
		return value;
	}

	private String getFirstParameterTypeTwo(String tag) {

		String firstParameter = null;
		int indexOfDot = tag.indexOf(".", 0);
		if (indexOfDot > 0) {
			firstParameter = tag.substring(0, indexOfDot);
		}
		return firstParameter;
	}

	private String getSecondParameterTypeTwo(String tag) {

		String secondParameter = null;
		int indexOfDot = tag.indexOf(".", 0);
		if (indexOfDot > 0) {
			secondParameter = tag.substring(indexOfDot + 1, tag.length());
			;
		}

		return secondParameter;
	}

}
