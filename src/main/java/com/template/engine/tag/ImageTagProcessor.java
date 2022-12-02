package com.template.engine.tag;

import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.template.engine.model.ImageDO;
import com.template.engine.utils.DocxConstants;
import com.template.engine.utils.DocxUtils;

public class ImageTagProcessor {

	public void fillImage(IBodyElement bodyElem, String tagText, Object mapValue)
			throws InvalidFormatException, IOException {

		fillImageValue(bodyElem, tagText, mapValue);
	}

	private void fillImageValue(IBodyElement bodyElem, String tagText, Object mapValue)
			throws InvalidFormatException, IOException {

		if (!DocxUtils.isNullEmpty(tagText)) {

			ImageDO image = (ImageDO) mapValue;

			if (bodyElem instanceof XWPFParagraph) {
				XWPFParagraph paragraph = (XWPFParagraph) bodyElem;
				XWPFRun run = paragraph.createRun();
				int imageFormat = getImageFormat(image);

				DocxUtils.replaceTextSegment(paragraph, tagText, DocxConstants.EMPTY_STRING);
				run.addPicture(image.getSourceStream(), imageFormat, image.getTitle(),
						Units.pixelToEMU(image.getWidth()), Units.pixelToEMU(image.getHeight()));

			} else if (bodyElem instanceof XWPFTable) {
				XWPFTable table = (XWPFTable) bodyElem;
				for (XWPFTableRow row : table.getRows()) {

					for (XWPFTableCell cell : row.getTableCells()) {

						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							XWPFRun run = paragraph.createRun();
							int imageFormat = getImageFormat(image);

							run.addPicture(image.getSourceStream(), imageFormat, image.getTitle(),
									Units.pixelToEMU(image.getWidth()), Units.pixelToEMU(image.getHeight()));
						}
					}
				}
			}
		}

	}

	/**
	 * Gets Format by specified content type.
	 *
	 * @param image image data
	 * @return image format or PNG if the content type is empty or unknown.
	 */
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
