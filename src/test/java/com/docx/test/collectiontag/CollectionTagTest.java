package com.docx.test.collectiontag;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.docx.test.model.Children;
import com.docx.test.model.Phone;
import com.docx.test.model.User;
import com.docx.test.utils.TestUtils;

public class CollectionTagTest {

	private static final String TEMPLATE_PATH = ".//datafiles//collection-tag-sample-template";
	private static final String OUTPUT_PATH = "C:\\Output-Docx-Engine";
	
	public static void main(String[] args) {
		try {
			System.out.println("Step 1 : Read template file");
			// Change template file name according to sample templates provided in datafiles folder
			String templateFileName = "collection-tag-in-paragraph-template-1.docx";
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

		Children firstSonAidid = new Children();
		firstSonAidid.setName("Amirul");
		firstSonAidid.setAge(5);
		Children secondSonAidid = new Children();
		secondSonAidid.setName("Iman");
		secondSonAidid.setAge(7);
		
		Children firstSonAmin = new Children();
		firstSonAmin.setName("Aisyah");
		Children secondSonAmin = new Children();
		secondSonAmin.setName("Azman");
		
		Children firstSonAfif = new Children();
		firstSonAfif.setName("Alisya");
		Children secondSonAfif = new Children();
		secondSonAfif.setName("Dayang");
		
		
		Phone phone1 = new Phone();
		phone1.setPhoneNo("012-XXXXXXXX");
		phone1.setProvider("U-Mobile");
		phone1.setDownlines(Arrays.asList(firstSonAidid, secondSonAidid));
		
		Phone phone2 = new Phone();
		phone2.setPhoneNo("012-XXXXXXXX");
		phone2.setProvider("Maxis");
		phone2.setDownlines(Arrays.asList(firstSonAmin, secondSonAmin));
		
		Phone phone3 = new Phone();
		phone3.setPhoneNo("019-XXXXXXXX");
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
		
		
		List<User> listOfUserCollection = new ArrayList<>(); //check type
		listOfUserCollection.add(aidid);
		listOfUserCollection.add(amin);
		listOfUserCollection.add(afif);
		
		List<Phone> listOfPhoneCollection = new ArrayList<>(); //check type
		listOfPhoneCollection.add(phone1);
		listOfPhoneCollection.add(phone2);
		listOfPhoneCollection.add(phone3);

		Map<String, Object> values = new HashMap<String, Object>();
		values.put("listOfUser", listOfUserCollection);
		values.put("listOfPhone", listOfPhoneCollection);

		return values;
	}
	
}
