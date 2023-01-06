package com.docx.tagprocessor;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.docx.model.ImageDO;
import com.docx.utils.DocxConstants;
import com.docx.utils.DocxUtils;

public class ImageTagProcessor {

	public void fillImage(IBodyElement bodyElem, String tagText, Object mapValue)
			throws InvalidFormatException, IOException {
		fillImageValue(bodyElem, tagText, mapValue);
	}

	private void fillImageValue(IBodyElement bodyElem, String tagText, Object mapValue)
			throws InvalidFormatException, IOException {

		ImageDO imageDO = (ImageDO) mapValue;

		XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
		XWPFRun run = paragraph.createRun();
		int imageFormat = getImageFormat(imageDO);

		DocxUtils.replaceTextSegment(paragraph, tagText, DocxConstants.EMPTY_STRING);
		run.addPicture(imageDO.getSourceStream(), imageFormat, imageDO.getTitle(),
				Units.pixelToEMU(imageDO.getWidth()), Units.pixelToEMU(imageDO.getHeight())); 
	}

	private int getImageFormat(ImageDO image) {
		if (image.getContentType() == null) {
			return Document.PICTURE_TYPE_PNG;
		}
		if (image.getContentType().contains("jpeg") || image.getContentType().contains("jpg")) {
			return Document.PICTURE_TYPE_JPEG;
		} else if (image.getContentType().contains("png")) {
			return Document.PICTURE_TYPE_PNG;
		} else if (image.getContentType().contains("tiff")) {
			return Document.PICTURE_TYPE_TIFF;
		} else if (image.getContentType().contains("bmp")) {
			return Document.PICTURE_TYPE_BMP;
		}

		// for all the rest formats
		return Document.PICTURE_TYPE_PNG;
	}

}
