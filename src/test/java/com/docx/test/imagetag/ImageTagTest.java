package com.docx.test.imagetag;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import javax.imageio.ImageIO;

import com.docx.model.ImageDO;
import com.docx.test.utils.TestUtils;

public class ImageTagTest {

	private static final String TEMPLATE_PATH = ".//datafiles//image-tag-sample-template";
	private static final String OUTPUT_PATH = "C:\\Output-Docx-Engine";
	
	public static void main(String[] args) {
		try {
			System.out.println("Step 1 : Read template file");
			String templateFileName = "image-tag-sample-template.docx";
			byte[] templateFile = TestUtils.readTemplateFile(templateFileName, TEMPLATE_PATH);

			System.out.println("Step 2 : Prepare sample value");
			Map<String, Object> values = sampleValues();

			System.out.println("Step 3 : Generate output file from template");
			byte[] outputFile = TestUtils.process(templateFile, values);
			String outputFileName = "sampleDocument.docx";
			TestUtils.writeFile(outputFile, outputFileName, OUTPUT_PATH);

			System.out.println("Completed.");
		} catch (Exception e) {
			System.err.format("ERROR : %s", e.getMessage());
		}
	}

	// Example value
	private static Map<String, Object> sampleValues() throws IOException {

		File imageFile = new File(".//datafiles//lightning.jpg");
		FileInputStream imageData = new FileInputStream(imageFile);
		BufferedImage image = ImageIO.read(imageFile);
		
		ImageDO imageDO = new ImageDO();
		imageDO.setTitle(imageFile.getName());
		imageDO.setSourceStream(imageData);
		imageDO.setContentType("jpg");
		imageDO.setWidth(image.getWidth());
		imageDO.setHeight(image.getHeight());

		Map<String, Object> values = new HashMap<String, Object>();
		values.put("lightning", imageDO);

		return values;
	}

}
