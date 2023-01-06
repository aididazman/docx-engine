package com.docx.utils;

import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTProofErr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import com.docx.model.TagDO;

public class DocxUtils {

	private DocxUtils() {
		throw new IllegalStateException("Util class");
	}

	private static final Pattern OBJECT_FIELD_PATTERN_1 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+:[a-zA-Z]+");
	private static final Pattern OBJECT_FIELD_PATTERN_2 = Pattern.compile("[a-zA-Z]+:[a-zA-Z]+");
	private static final Pattern OBJECT_FIELD_PATTERN_3 = Pattern.compile("[a-zA-Z]+\\.[a-zA-Z]+");

	static String getFirstParameter(String tagText) {

		String objectName = null;
		int indexOfColon = tagText.indexOf(":", 0);
		if (indexOfColon > 0) {
			objectName = tagText.substring(0, indexOfColon);
		}

		return objectName;
	}

	static String getSecondParameter(String tagText) {

		String objectField = null;
		int indexOfColon = tagText.indexOf(":", 0);
		if (indexOfColon > 0) {
			objectField = tagText.substring(indexOfColon + 1, tagText.length());
			;
		}

		return objectField;
	}

	static String getFirstParameterTypeTwo(String tagText) {

		String objectName = null;
		int indexOfDot = tagText.indexOf(".", 0);
		if (indexOfDot > 0) {
			objectName = tagText.substring(0, indexOfDot);
		}
		return objectName;
	}

	static String getSecondParameterTypeTwo(String tagText) {

		String objectField = null;
		int indexOfDot = tagText.indexOf(".", 0);
		if (indexOfDot > 0) {
			objectField = tagText.substring(indexOfDot + 1, tagText.length());
		}

		return objectField;
	}

	public static String getObjectName(String tagText) {

		String objectName = null;
		if (OBJECT_FIELD_PATTERN_1.matcher(tagText).matches()) {
			objectName = getFirstParameter(tagText);
			objectName = getFirstParameterTypeTwo(objectName);
		} else if (OBJECT_FIELD_PATTERN_2.matcher(tagText).matches()) {
			objectName = getFirstParameter(tagText);
		} else if (OBJECT_FIELD_PATTERN_3.matcher(tagText).matches()) {
			objectName = getFirstParameterTypeTwo(tagText);
		}
		return objectName;
	}

	public static String getObjectField(String tagText) {

		String objectField = null;
		if (OBJECT_FIELD_PATTERN_1.matcher(tagText).matches()
				|| OBJECT_FIELD_PATTERN_2.matcher(tagText).matches()) {
			objectField = getSecondParameter(tagText);
		} else if (OBJECT_FIELD_PATTERN_3.matcher(tagText).matches()) {
			objectField = getSecondParameterTypeTwo(tagText);
		}
		return objectField;
	}

	public static String addTagBracket(String tagValue) {

		StringBuilder tagValueWithBracket = new StringBuilder(tagValue);
		tagValueWithBracket.insert(tagValue.length(), DocxConstants.DEFAULT_TAG_END);
		tagValueWithBracket.insert(0, DocxConstants.DEFAULT_TAG_START);

		return tagValueWithBracket.toString();
	}

	public static List<TagDO> getTagsByElement(String elementText, int tagStartOffset, List<TagDO> tagDOList,
			IBodyElement tagElement) throws Exception {

		tagStartOffset = elementText.indexOf(DocxConstants.DEFAULT_TAG_START, tagStartOffset);
		if (tagStartOffset >= 0) {
			int tagEndOffset = elementText.indexOf(DocxConstants.DEFAULT_TAG_END, tagStartOffset);

			if (tagEndOffset < 0) {
				throw new Exception("No closing tag found for line " + elementText);
			}

			String tagText = elementText.substring(tagStartOffset + DocxConstants.DEFAULT_TAG_START.length(),
					tagEndOffset);

			TagDO tagDO = new TagDO(tagText, tagEndOffset, tagElement);
			tagDOList.add(tagDO);

			// recursive to proceed to get other tags in the same text
			getTagsByElement(elementText, tagEndOffset, tagDOList, tagElement);
		}

		return tagDOList;
	}

	public static String getFieldValue(String tagObjectField, Object mapValue) {

		Object value = null;
		try {
			value = PropertyUtils.getSimpleProperty(mapValue, tagObjectField);
		} catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
			throw new RuntimeException("Cannot get tag " + tagObjectField + " value from the context");
		}
		String tagValue = value == null ? null : value.toString(); 
		if (tagValue == null) {
			tagValue = DocxConstants.EMPTY_STRING;
		}

		return tagValue;
	}

	public static String processValue(String tagName, Map<String, Object> resolutionAttributesMap) {

		String value = null;
		Object mapValue = null;
		String tagObjectField = null;

		if (OBJECT_FIELD_PATTERN_3.matcher(tagName).matches()) {
			mapValue = resolutionAttributesMap.get(tagName);

			if (mapValue != null) {
				tagObjectField = getSecondParameterTypeTwo(tagName); // get name from user.name
				if (!isNullEmpty(tagObjectField)) {
					value = getFieldValue(tagObjectField, mapValue);
				}
			}
		}

		else {
			mapValue = resolutionAttributesMap.get(tagName);
			if (mapValue != null) {
				value = mapValue.toString();
			}
		}
		return value;
	}

	public static String getMapKey(String tagName) {

		String firstParameter = null;
		firstParameter = getFirstParameter(tagName);
		return firstParameter;
	}

	public static IBodyElement getNextElement(IBodyElement elem) {
		for (int i = 0; i < elem.getBody().getBodyElements().size() - 1; i++) {
			if (elem.getBody().getBodyElements().get(i) == elem) {
				return elem.getBody().getBodyElements().get(i + 1);
			}
		}

		return null;
	}

	public static int getElementIndex(IBodyElement element) {
		IBody body = element.getBody();
		for (int i = 0; i < body.getBodyElements().size(); i++) {
			if (body.getBodyElements().get(i) == element) {
				return i;
			}
		}

		return -1;
	}

	public static int getParagraphIndex(List<XWPFParagraph> paragraphList, XWPFParagraph paragraph) {
		int index = -1;
		for (XWPFParagraph paraprgaph : paragraphList) {
			if (paraprgaph instanceof XWPFParagraph) {
				index++;
			}
			if (paraprgaph == paragraph) {
				return index;
			}
		}
		return -1;
	}

	public static String getTagName(String tagText, String tagPrefix) {
		return tagText.substring(tagPrefix.length());
	}
	
	
	/**
	 * @author Axel Richter
	 * @see https://stackoverflow.com/questions/65275097/apache-poi-my-placeholder-is-treated-as-three-different-runs
	 */ 
	public static void replaceTextSegment(XWPFParagraph paragraph, String textToFind, String replacement) {
		TextSegment foundTextSegment = null;
		PositionInParagraph startPos = new PositionInParagraph(0, 0, 0);
		while ((foundTextSegment = searchText(paragraph, textToFind, startPos)) != null) { // search all text segments
																							// having text to find

			System.out.println(foundTextSegment.getBeginRun() + ":" + foundTextSegment.getBeginText() + ":"
					+ foundTextSegment.getBeginChar());
			System.out.println(foundTextSegment.getEndRun() + ":" + foundTextSegment.getEndText() + ":"
					+ foundTextSegment.getEndChar());

			// maybe there is text before textToFind in begin run
			XWPFRun beginRun = paragraph.getRuns().get(foundTextSegment.getBeginRun());		
			String textInBeginRun = beginRun.getText(foundTextSegment.getBeginText());
			String textBefore = textInBeginRun.substring(0, foundTextSegment.getBeginChar()); // we only need the text
																								// before

			// maybe there is text after textToFind in end run
			XWPFRun endRun = paragraph.getRuns().get(foundTextSegment.getEndRun());
			String textInEndRun = endRun.getText(foundTextSegment.getEndText());
			String textAfter = textInEndRun.substring(foundTextSegment.getEndChar() + 1); // we only need the text after

			if (foundTextSegment.getEndRun() == foundTextSegment.getBeginRun()) {
				textInBeginRun = textBefore + replacement + textAfter; // if we have only one run, we need the text
																		// before, then the replacement, then the text
																		// after in that run
			} else {
				textInBeginRun = textBefore + replacement; // else we need the text before followed by the replacement
															// in begin run
				endRun.setText(textAfter, foundTextSegment.getEndText()); // and the text after in end run
			}

			beginRun.setText(textInBeginRun, foundTextSegment.getBeginText());

			// runs between begin run and end run needs to be removed
			for (int runBetween = foundTextSegment.getEndRun() - 1; runBetween > foundTextSegment
					.getBeginRun(); runBetween--) {
				paragraph.removeRun(runBetween); // remove not needed runs
			}

		}
	}

	/**
	 * this methods parse the paragraph and search for the string searched. If it
	 * finds the string, it will return true and the position of the String will be
	 * saved in the parameter startPos.
	 *
	 * @param searched
	 * @param startPos
	 */
	static TextSegment searchText(XWPFParagraph paragraph, String searched, PositionInParagraph startPos) {
		int startRun = startPos.getRun(), startText = startPos.getText(), startChar = startPos.getChar();
		int beginRunPos = 0, candCharPos = 0;
		boolean newList = false;

		// CTR[] rArray = paragraph.getRArray(); //This does not contain all runs. It
		// lacks hyperlink runs for ex.
		List<XWPFRun> runs = paragraph.getRuns();

		int beginTextPos = 0, beginCharPos = 0; // must be outside the for loop

		// for (int runPos = startRun; runPos < rArray.length; runPos++) {
		for (int runPos = startRun; runPos < runs.size(); runPos++) {
			// int beginTextPos = 0, beginCharPos = 0, textPos = 0, charPos; //int
			// beginTextPos = 0, beginCharPos = 0 must be outside the for loop
			int textPos = 0, charPos;
			// CTR ctRun = rArray[runPos];
			CTR ctRun = runs.get(runPos).getCTR();
			XmlCursor c = ctRun.newCursor();
			c.selectPath("./*");
			try {
				while (c.toNextSelection()) {
					XmlObject o = c.getObject();
					if (o instanceof CTText) {
						if (textPos >= startText) {
							String candidate = ((CTText) o).getStringValue();
							if (runPos == startRun) {
								charPos = startChar;
							} else {
								charPos = 0;
							}

							for (; charPos < candidate.length(); charPos++) {
								if ((candidate.charAt(charPos) == searched.charAt(0)) && (candCharPos == 0)) {
									beginTextPos = textPos;
									beginCharPos = charPos;
									beginRunPos = runPos;
									newList = true;
								}
								if (candidate.charAt(charPos) == searched.charAt(candCharPos)) {
									if (candCharPos + 1 < searched.length()) {
										candCharPos++;
									} else if (newList) {
										TextSegment segment = new TextSegment();
										segment.setBeginRun(beginRunPos);
										segment.setBeginText(beginTextPos);
										segment.setBeginChar(beginCharPos);
										segment.setEndRun(runPos);
										segment.setEndText(textPos);
										segment.setEndChar(charPos);
										return segment;
									}
								} else {
									candCharPos = 0;
								}
							}
						}
						textPos++;
					} else if (o instanceof CTProofErr) {
						c.removeXml();
					} else if (o instanceof CTRPr) {
						// do nothing
					} else {
						candCharPos = 0;
					}
				}
			} finally {
				c.dispose();
			}
		}
		return null;
	}

	public static boolean isNullEmpty(String inputStr) {
		return (inputStr == null || "".equals(inputStr));
	}

	public static <T> boolean isNullEmpty(List<T> list) {
		return (list == null || list.isEmpty());
	}
}
