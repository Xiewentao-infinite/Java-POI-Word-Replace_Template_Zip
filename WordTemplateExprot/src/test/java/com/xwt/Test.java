package com.xwt;


import com.xwt.word.WordReplace;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Test {
	@org.junit.Test
	public void method() throws IOException {
		WordReplace wordReplace = new WordReplace();
		wordReplace.wordTrans();
	}
	@org.junit.Test
	public void  makeZip() throws IOException {
		FileOutputStream fileOutputStream = new FileOutputStream("src/main/resources/template/Words.zip");
		File file = new File("src/main/resources/template/Template3.docx");
		FileInputStream fileInputStream = new FileInputStream(file);

		ZipOutputStream zip = new ZipOutputStream(fileOutputStream);
		zip.putNextEntry(new ZipEntry( "Template3.docx"));

		byte bytes[]=new byte[1024*5];
		int len;
		while ((len = fileInputStream.read(bytes)) != -1) {
			zip.write(bytes, 0, len);
		}

		zip.flush();
		zip.closeEntry();

	}

}
