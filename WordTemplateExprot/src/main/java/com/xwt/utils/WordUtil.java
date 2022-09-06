package com.xwt.utils;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @author Xie Wentao
 */
public class WordUtil {

	public static void replaceWordContent(XWPFDocument document, Map<String, String> relationMap, String url) throws Exception {
		List<XWPFTable> tables = document.getTables();

		Map<String, String> textMap = new HashMap<String, String>();
		String date = DateUtils.getDate();
		String[] split = date.split("-");
		String year = split[0];
		String month = split[1];
		String day = split[2];
		textMap.put("{year}",year+"年"+month+"月"+day+"日");
		replaceText(document,textMap);
		for (XWPFTable table : tables) {
			// 获取表格的每一行数据
			for (XWPFTableRow row : table.getRows()) {
				// 获取表格的每一个单元格
				for (XWPFTableCell cell : row.getTableCells()) {
					String text = cell.getText();
					// 内容换行
					if (!isEmpty(cell.getText()) && cell.getText().contains("\n")) {
						for (XWPFParagraph paragraph : cell.getParagraphs()) {
							for (XWPFRun run : paragraph.getRuns()) {
								if (run.getText(0) != null && run.getText(0).contains("\n")) {
									String[] lines = run.getText(0).split("\n");
									if (lines.length > 0) {
										// set first line into XWPFRun
										run.setText(lines[0], 0);
										for (int i = 1; i < lines.length; i++) {
											// add break and insert new text
											run.addBreak();
											run.setText(lines[i]);
										}
									}
								}
							}
						}
					}

					// 再新增一个格式
					// XWPFParagraph paragraph = cell.addParagraph();
					// 设置行间距
					// paragraph.setSpacingBetween(1);
					if (!text.equals("{zhaopian}"))
					{
						// 移除操作, 因为poi底层对word是追加操作, 如果不移除, 会在原内容是拼接,
						cell.removeParagraph(0);
						// 替换后的新值
						String newText = replaceCellContent(text, relationMap);
						// 将值放入单元格
						cell.setText(newText);
					}
					replacePhoto(document,url);
				}
			}
		}
	}

	public static String replaceCellContent(String cellContent, Map<String, String> relationMap) throws Exception {

		if (isEmpty(cellContent))
		{
			return "";
		}
		for (Map.Entry<String, String> entry : relationMap.entrySet()) {
			String k = entry.getKey();
			String v = entry.getValue();
			if (!k.equals("{zhaopian}")) {
				if (cellContent.contains(k)) {
					cellContent = cellContent.replace(k, v);
					return cellContent;
				}
			}
		}

		return "";
	}

	public static boolean isEmpty(String str)
	{
		if (null == str || str.length() == 0)
		{
			return true;
		}
		return false;
	}


	/**
	 * 替换非表格埋点值
	 *
	 * @param document
	 * @param textMap  需要替换的文本入参
	 */
	public static void replaceText(XWPFDocument document, Map<String, String> textMap) {
		List<XWPFParagraph> paras = document.getParagraphs();
		Set<String> keySet = textMap.keySet();
		for (XWPFParagraph para : paras) {
			//当前段落的属性
//			String str = para.getText();
			List<XWPFRun> list = para.getRuns();
			for (XWPFRun run : list) {
				for (String key : keySet) {
					if (key.equals(run.text())) {
						run.setText(textMap.get(key), 0);
					}
				}
			}

		}
	}
	public static void replacePhoto(XWPFDocument document,String url)throws Exception {

		File frontImg = new File(url);

		FileInputStream fileInputStreamFront = new FileInputStream(frontImg);
		List<XWPFTable> tables = document.getTables();
		HashMap<String, FileInputStream> map = new HashMap<>();
		map.put("{zhaopian}", fileInputStreamFront);
		insertImg(map, tables,url);
		fileInputStreamFront.close();
	}

	public  static  void   insertImg(HashMap<String,FileInputStream> map, List<XWPFTable> tables, String url){
		//这个循环可以去掉，因为模板中的table是可以通过id直接获取的
		for (XWPFTable table:tables) {
			List<XWPFTableRow> rows = table.getRows();
			for (XWPFTableRow row:rows){
				List<XWPFTableCell> tableCells = row.getTableCells();
				for (XWPFTableCell cell:tableCells) {
					//这个map的循环也可以去掉，把替换的变量写死即可，里面的if就得是“变量名”.equals("")。
					for(Map.Entry<String, FileInputStream> entry : map.entrySet()){

						if(entry.getKey().equals(cell.getText())){
							List<XWPFParagraph> paragraphs = cell.getParagraphs();
							for (XWPFParagraph paragraph :paragraphs) {
								List<XWPFRun> runs = paragraph.getRuns();

								for (XWPFRun run:runs){
									try {
										run.addPicture(entry.getValue(), XWPFDocument.PICTURE_TYPE_JPEG, url, Units.toEMU(90), Units.toEMU(130));
										run.setText("  ", 0);
									}catch (Exception e){
										e.printStackTrace();
									}
								}
							}
							cell.setText(" ");
						}
					}

				}
			}
		}
	}

}
