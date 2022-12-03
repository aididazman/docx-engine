package com.template.engine.tag;

import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import com.template.engine.model.TagInfo;
import com.template.engine.utils.DocxUtils;

public class HeaderFooterTagProcessor {
	
	private static final Pattern FIELD_OBJECT_PATTERN = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+");

	public String processValue(TagInfo tag, Map<String, Object> resolutionAttributesMap, String tagName) {

		String value = null;
		Object mapValue = null;
		// get tag text from tag, example field:name / field:user.name -> name /
		// user.name
		
		// check whether tag text match type pattern
		if (FIELD_OBJECT_PATTERN.matcher(tagName).matches()) {
			String tagObjectName = getFirstParameterTypeTwo(tagName);// get user from user.name
			mapValue = resolutionAttributesMap.get(tagName);
			// check whether tag object name equals the name of the class from map value
			if (mapValue.getClass().getSimpleName().equalsIgnoreCase(tagObjectName)) {
				String tagObjectField = getSecondParameterTypeTwo(tagName); // get name from user.name
				value = DocxUtils.getFieldValue(tagObjectField, mapValue); 
			}
		}

		else {
			mapValue = resolutionAttributesMap.get(tagName);
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

	public void fillHeaderFooterTag(XWPFParagraph paragraph, String textToBeReplaced, String textValue) {
		DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);
	}

}
