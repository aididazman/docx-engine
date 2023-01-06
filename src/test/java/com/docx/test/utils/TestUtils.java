package com.docx.test.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.docx.impl.DocxEngine;

public class TestUtils {

	private TestUtils() {
		throw new IllegalStateException("Util class");
	}
	
	public static byte[] readTemplateFile(String templateFileName, String templatePath) throws Exception {
		if (templateFileName.equals(null))
			return null;

		File templateFile = new File(String.format("%s\\%s", templatePath, templateFileName));
		if (!templateFile.exists())
			throw new Exception(String.format("Template file not found. Template path %s", templateFile.getPath()));

		return Files.readAllBytes(templateFile.toPath());
	}
	
	public static void writeFile(byte[] fileContent, String outputFileName, String outputPath) throws Exception {
		if (outputFileName.equals(null))
			throw new Exception("Output file name is empty.");

		if (fileContent == null)
			throw new Exception("File content is empty.");

		File outputFile = new File(String.format("%s\\%s", outputPath, outputFileName));
		try {
			if (!outputFile.getParentFile().exists())
				outputFile.getParentFile().mkdirs();

			outputFile.createNewFile();

			OutputStream os = new FileOutputStream(outputFile);
			os.write(fileContent);
			os.close();
		} catch (Exception e) {
			throw new Exception(String.format("Failed generate ouput file. Errror : %s. Path : %s", e.getMessage(),
					outputFile.getPath()));
		}
	}
	
	public static byte[] process(byte[] content, Map<String, Object> resolutionAttributesMap) {

		try (InputStream inputStream = new ByteArrayInputStream(content);
				XWPFDocument document = new XWPFDocument(inputStream);
				ByteArrayOutputStream out = new ByteArrayOutputStream()) {

			if (content == null)
				throw new Exception("Template content is null.");
			
			DocxEngine docxEngine = new DocxEngine(content, resolutionAttributesMap);
			return docxEngine.generateDocument();

		} catch (Exception e) {
			e.printStackTrace();
		}

		return null;

	}
	
}
