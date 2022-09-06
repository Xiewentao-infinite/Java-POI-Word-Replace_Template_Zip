package com.xwt.word;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

import static com.xwt.utils.WordUtil.replaceWordContent;

public class WordReplace {

	public void wordTrans() throws IOException {
		FileInputStream inputStream = null;
		XWPFDocument document = null;
		FileOutputStream outputStream = new FileOutputStream( "src/main/resources/template/Template3.docx");
		try {
			inputStream = new FileInputStream("src/main/resources/template/Template.docx");
			document = new XWPFDocument(inputStream);
			Map<String, String> map = new HashMap<>();
			map.put("{nation}", "Rnation");
			map.put("{health}", "Rhealth");
			map.put("{name}", "Rname");
			map.put("{nativePlaceName}", "RnativePlaceName");
			map.put("{birthPlaceName}", "RbirthPlaceName");
			map.put("{partyTime}", "RpartTime");
			map.put("{workTime}", "RworkTime");
			map.put("{gender}", "Man");
			map.put("{WorkExperience}",  "\r\n" +"2020-01-01 in google;" + "\r\n" + "2021-01-01 in facebook;");

			// replacePhoto
			replaceWordContent(document, map, "src/main/resources/template/photo.png");
			for (String key : map.keySet()) {
				System.out.println("Key: " + key + ", Value: " + map.get(key));
			}
			// make new word
			document.write(outputStream);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(inputStream);
			IOUtils.closeQuietly(outputStream);
			IOUtils.closeQuietly(document);
		}
	}
}
