package com.template.engine.tag;

import java.lang.reflect.InvocationTargetException;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import com.template.engine.model.TagInfo;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class FooterTagProcessor {

	public String getValue(TagInfo tag, Object mapValue) {
		String tagValue = getTagValue(tag, mapValue);
		return tagValue;
	}

	private String getTagValue(TagInfo tag, Object mapValue) {
		Object value = null;
		try {
			value = PropertyUtils.getSimpleProperty(mapValue,
					DocxUtils.getTagName(tag, DocxConstants.TAG_PREFIX_FOOTER));
		} catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
			throw new RuntimeException("Cannot get tag " + tag + " value from the context");
		}
		String tagValue = value == null ? null : value.toString();
		if (tagValue == null) {
			tagValue = DocxConstants.EMPTY_STRING;
		}

		return tagValue;
	}

	public void fillFooterTag(XWPFParagraph paragraph, String textToBeReplaced, String textValue) {
		DocxUtils.replaceTextSegment(paragraph, textToBeReplaced, textValue);	
	}

}
