package com.docx.test.fieldtag;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import com.docx.test.model.HeaderAndFooter;
import com.docx.test.model.User;
import com.docx.test.utils.TestUtils;

public class FieldTagTest {

	private static final String TEMPLATE_PATH = ".//datafiles//field-tag-sample-template";
	private static final String OUTPUT_PATH = "C:\\Output-Docx-Engine";
	
	public static void main(String[] args) {
		try {
			System.out.println("Step 1 : Read template file");
			String templateFileName = "field-tag-sample-template.docx";
			byte[] templateFile = TestUtils.readTemplateFile(templateFileName, TEMPLATE_PATH);

			System.out.println("Step 2 : Prepare sample value");
			Map<String, Object> values = sampleValues();

			System.out.println("Step 3 : Generate output file from template");
			byte[] outputFile = TestUtils.process(templateFile, values);
			String outputFileName = "sample-document.docx";
			TestUtils.writeFile(outputFile, outputFileName, OUTPUT_PATH);

			System.out.println("Completed.");
		} catch (Exception e) {
			System.err.format("ERROR : %s", e.getMessage());
		}
	}

	// Example value
	private static Map<String, Object> sampleValues() throws IOException {
		
		User user = new User();
		user.setName("Aidid");
		user.setAge(25);
		
		HeaderAndFooter headerAndFooter = new HeaderAndFooter();
		headerAndFooter.setHeader("This is header");
		
		Map<String, Object> values = new HashMap<String, Object>();
		values.put("headerAndFooter.header", headerAndFooter);
		values.put("footer", "This is footer");
		values.put("firstName", "Muhammad");
		values.put("lastName", "Amin");
		values.put("user.name", user);
		values.put("user.age", user);

		return values;
	}

}
