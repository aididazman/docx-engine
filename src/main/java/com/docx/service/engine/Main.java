package com.docx.service.engine;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.docx.service.model.Children;
import com.docx.service.model.HeaderAndFooter;
import com.docx.service.model.ImageDO;
import com.docx.service.model.ListOfPhone;
import com.docx.service.model.ListOfUser;
import com.docx.service.model.Phone;
import com.docx.service.model.User;
import com.docx.service.templatemode.TemplateMode;
import com.google.common.io.ByteSource;

public class Main {

	private static final String TEMPLATE_PATH = "C:\\Users\\aidid\\Desktop";
	private static final String OUTPUT_PATH = "C:\\Users\\aidid\\Desktop\\output";

	public static void main(String[] args) {
		try {
			System.out.println("Step 1 : Read template file");
			String templateFileName = "Sample template.docx";
			byte[] templateFile = readTemplateFile(templateFileName);

			System.out.println("Step 2 : Prepare sample value");
			Map<String, Object> values = sampleValues();

			System.out.println("Step 3 : Generate output file from template");
			byte[] outputFile = process(templateFile, null, TemplateMode.DOCX, values);
			String outputFileName = "sampleDocument.docx";
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

		Children firstSonAidid = new Children();
		firstSonAidid.setName("Amirul");
		Children secondSonAidid = new Children();
		secondSonAidid.setName("Iman");
		
		Children firstSonAmin = new Children();
		firstSonAmin.setName("Aisyah");
		Children secondSonAmin = new Children();
		secondSonAmin.setName("Azman");
		
		Children firstSonAfif = new Children();
		firstSonAfif.setName("Alisya");
		Children secondSonAfif = new Children();
		secondSonAfif.setName("Dayang");
		
		
		Phone phone1 = new Phone();
		phone1.setPhoneNo("012-3103004");
		phone1.setProvider("U-Mobile");
		phone1.setDownlines(Arrays.asList(firstSonAidid, secondSonAidid));
		
		Phone phone2 = new Phone();
		phone2.setPhoneNo("012-3456789");
		phone2.setProvider("Maxis");
		phone2.setDownlines(Arrays.asList(firstSonAmin, secondSonAmin));
		
		Phone phone3 = new Phone();
		phone3.setPhoneNo("019-8765432");
		phone3.setProvider("Celcom");
		phone3.setDownlines(Arrays.asList(firstSonAfif, secondSonAfif));
		
		Phone phone4 = new Phone();
		phone4.setPhoneNo("019-XXXXXXXX");
		phone4.setProvider("TEST TELCO");
		phone4.setDownlines(Arrays.asList(firstSonAfif, secondSonAfif));

		User aidid = new User();
		aidid.setName("Aidid");
		aidid.setAge(24);
		aidid.setPhones(Arrays.asList(phone1, phone4));
		aidid.setChilds(Arrays.asList(firstSonAidid, secondSonAidid));
		
		User amin = new User();
		amin.setName("Amin");
		amin.setAge(30);
		amin.setPhones(Arrays.asList(phone2, phone4));
		amin.setChilds(Arrays.asList(firstSonAmin, secondSonAmin));
		
		User afif = new User();
		afif.setName("Afif");
		afif.setAge(29);
		afif.setPhones(Arrays.asList(phone3, phone4));
		afif.setChilds(Arrays.asList(firstSonAfif, secondSonAfif));

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
		
		List<User> listOfUserCollection = new ArrayList<>(); //check type
		listOfUserCollection.add(aidid);
		listOfUserCollection.add(amin);
		listOfUserCollection.add(afif);
		
		List<Phone> listOfPhoneCollection = new ArrayList<>(); //check type
		listOfPhoneCollection.add(phone1);
		listOfPhoneCollection.add(phone2);
		listOfPhoneCollection.add(phone3);
		
		
		Map<String, Object> values = new HashMap<String, Object>();
		values.put("listOfPhone", listOfPhoneCollection);
		//values.put("user.phones", aidid); // for non nested 
		values.put("listOfUser", listOfUserCollection);
		//values.put("listOfUser", listOfUser);
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
			
			DocxEngine docxEngine = new DocxEngine(content, resolutionAttributesMap);
			return docxEngine.generateDocument();

		} catch (Exception e) {
			e.printStackTrace();
		}

		return null;

	}

}
