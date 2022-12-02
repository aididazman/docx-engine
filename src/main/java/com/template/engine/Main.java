package com.template.engine;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.google.common.io.ByteSource;
import com.template.engine.model.HeaderAndFooter;
import com.template.engine.model.ImageDO;
import com.template.engine.model.ListOfPhone;
import com.template.engine.model.ListOfUser;
import com.template.engine.model.Phone;
import com.template.engine.model.User;
import com.template.engine.tag.TestProcessor;
import com.template.engine.templatemode.TemplateMode;

public class Main {

	private static final String TEMPLATE_PATH = "C:\\Users\\aidid\\Desktop";
	private static final String OUTPUT_PATH = "C:\\Users\\aidid\\Desktop\\output";

	public static void main(String[] args) {
		try {
			System.out.println("Step 1 : Read template file");
			String templateFileName = "nestedCollection-2.docx";
			byte[] templateFile = readTemplateFile(templateFileName);

			System.out.println("Step 2 : Prepare sample value");
			Map<String, Object> values = sampleValues();

			System.out.println("Step 3 : Generate output file from template");
			byte[] outputFile = process(templateFile, null, TemplateMode.DOCX, values);
			String outputFileName = "nestedCollection_2.docx";
			writeFile(outputFile, outputFileName);

			System.out.println("Completed.");
		} catch (Exception e) {
			System.err.format("ERROR : %s", e.getMessage());
		}
	}

	protected static byte[] readTemplateFile(String templateFileName) throws Exception {
		if (templateFileName.equals(null))
			return null;

		File templateFile = new File(String.format("%s\\%s", TEMPLATE_PATH, templateFileName));
		if (!templateFile.exists())
			throw new Exception(String.format("Template file not found. Template path %s", templateFile.getPath()));

		return Files.readAllBytes(templateFile.toPath());
	}

	private static void writeFile(byte[] fileContent, String outputFileName) throws Exception {
		if (outputFileName.equals(null))
			throw new Exception("Output file name is empty.");

		if (fileContent == null)
			throw new Exception("File content is empty.");

		File outputFile = new File(String.format("%s\\%s", OUTPUT_PATH, outputFileName));
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

	// Example value
	private static Map<String, Object> sampleValues() throws IOException {

		Phone phone1 = new Phone();
		phone1.setPhoneNo("012-3103004");
		phone1.setProvider("U-Mobile");
		
		Phone phone2 = new Phone();
		phone2.setPhoneNo("012-3456789");
		phone2.setProvider("Maxis");
		
		Phone phone3 = new Phone();
		phone3.setPhoneNo("019-8765432");
		phone3.setProvider("Celcom");
		
		Phone phone4 = new Phone();
		phone4.setPhoneNo("019-XXXXXXXX");
		phone4.setProvider("TEST TELCO");

		User aidid = new User();
		aidid.setName("Aidid");
		aidid.setAge(24);
		aidid.setPhones(Arrays.asList(phone1, phone4));
		
		User amin = new User();
		amin.setName("Amin");
		amin.setAge(30);
		amin.setPhones(Arrays.asList(phone2, phone4));
		
		User afif = new User();
		afif.setName("Afif");
		afif.setAge(29);
		afif.setPhones(Arrays.asList(phone3, phone4));

		File imageFile = new File(".//datafiles//lightning.jpg");
		FileInputStream imageData = new FileInputStream(imageFile);
		BufferedImage picture = ImageIO.read(imageFile);
		
		int width = picture.getWidth();
		int height = picture.getHeight();
		int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
		String imageFileName = imageFile.getName();
		
		ImageDO image = new ImageDO();
		image.setTitle(imageFileName);
		image.setSourceStream(imageData);
		image.setContentType("jpg");
		image.setWidth(width);
		image.setHeight(height);
		
		HeaderAndFooter headerAndFooter  = new HeaderAndFooter();
		headerAndFooter.setHeader("THIS IS HEADER");
		headerAndFooter.setFooter("THIS IS FOOTER");
		
		ListOfUser listOfUser = new ListOfUser();
		listOfUser.setUsers(Arrays.asList(aidid, amin, afif));
		
		ListOfPhone listOfPhone = new ListOfPhone();
		listOfPhone.setPhones(Arrays.asList(phone1, phone2, phone3));
		
		Map<String, Object> values = new HashMap<String, Object>();
		values.put("listOfPhone", listOfPhone);
		//values.put("user.phones", aidid);
		values.put("listOfUser", listOfUser);
		values.put("headerAndFooter.header", headerAndFooter);
		values.put("footer", "This is footer");
		values.put("lightning", image);
		values.put("user.age", aidid);
		values.put("user.name", aidid);
		values.put("firstName", "Muhammad");
		values.put("lastName", "Amin");

		return values;

	}

	private static byte[] process(byte[] content, String encoding, TemplateMode templateMode,
			Map<String, Object> resolutionAttributesMap) {

		try (InputStream inputStream = ByteSource.wrap(content).openStream();
				XWPFDocument document = new XWPFDocument(inputStream);
				ByteArrayOutputStream out = new ByteArrayOutputStream()) {

			if (content == null)
				throw new Exception("Template content is null.");

//			AllTagProcessor allTagProcessor = new AllTagProcessor(content, resolutionAttributesMap);
//			return allTagProcessor.generateDocument();
			
			TestProcessor testTagProcessor = new TestProcessor(content, resolutionAttributesMap);
			return testTagProcessor.generateDocument();

		} catch (Exception e) {
//			System.err.format("ERROR : %s", e.getMessage());
			e.printStackTrace();
		}

		return null;

	}

}