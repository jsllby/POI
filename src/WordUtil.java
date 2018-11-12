package com.ruite.util.poi;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class WordUtil
{

	public static final String IMAGE_ = "IMAGE_";
	private static final String EL_LEFT = "{";
	private static final String EL_RIGHT = "}";
	private static final double WIDTH_MAX_SIZE = 409.0;			// A4纸尺寸
	private static final double HEIGHT_MAX_SIZE = 681.0;
	private XWPFDocument document = null;
	private InputStream is = null;								// 用于关闭输入流
	
	private int maxTableRowCount = 30;
	
	private List<String> noSNTable = null;

	/**
	 * Read and write word by template
	 * 
	 * @param templatePath
	 *            The path of the word template
	 * @param outputPath
	 *            The path of the result word file to write out
	 * @param map
	 *            The rule that is how to read and write the word for example:
	 *            e.g string value replace: The map's key is the text which is
	 *            ready to repalce in the word and value of the map is a text
	 *            insert into word. e.g image insert: The map's key must start
	 *            with 'IMAGE_' and value of the map must be put a path of a
	 *            image. e.g list value insert: The map's key is the name of a
	 *            table in word. The value of the map must be a instance of
	 *            calss List and the members of list is a instance of calss Map
	 *            and this map as the string value replace map
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	public boolean generateWordFromTemplate(String templatePath, Map<String, Object> map) throws Exception
	{
		//document = new XWPFDocument(POIXMLDocument.openPackage(templatePath));
		is = new FileInputStream(templatePath);
		return generateWordFromTemplate(is, map);
	}

	/**
	 * Read and write word by template in inputStream
	 * 替换基本内容
	 * 
	 * @param inputStream
	 * @param outputPath
	 * @param map
	 * @return
	 * @throws InvalidFormatException
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public boolean generateWordFromTemplate(java.io.InputStream inputStream, Map<String, Object> map) throws Exception
	{
		document = new XWPFDocument(inputStream);
		return readWordDocument(map);
	}

	private boolean readWordDocument(Map<String, Object> map) throws Exception, IOException
	{
		// replace the text of header
		if (document.getHeaderList() != null && document.getHeaderList().size() > 0)
		{
			XWPFHeader header = document.getHeaderArray(0);
			List<XWPFParagraph> listHeader = header.getParagraphs();
			for (XWPFParagraph paragraph : listHeader)
			{
				replaceWordText(paragraph, map);
			}
		}
		// replace the text of body
		Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
		while (itPara.hasNext())
		{
			XWPFParagraph paragraph = itPara.next();
//			replaceWordText(paragraph, map);
			replaceWordTextJust(paragraph, map);
			// wordReadRule.addPictureInWord(document);
		}
		// replace the text of table
		Iterator<XWPFTable> itTable = document.getTablesIterator();
		while (itTable.hasNext())
		{
			XWPFTable table = itTable.next();
//			readWordTableRangeRule(table, map);
			replaceTableWordTextJust(table, map);
		}
		
		return true;
	}

	/**
	 * replace text and images
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void replaceWordText(XWPFParagraph para, Map<String, Object> params) throws Exception
	{
		List<XWPFRun> runs;
		Matcher matcher;
		
		if (this.matcher(para.getParagraphText()).find())
		{
			runs = para.getRuns();
			for (int i = 0; i < runs.size(); i++)
			{
				XWPFRun run = runs.get(i);
				String runText = run.toString();
				matcher = this.matcher(runText);

				if (matcher.find())
				{
					// 替换1对1变量
					if (params.get(matcher.group(1)) instanceof String)
					{
						String str = (String) params.get(matcher.group(1));
						runText = matcher.replaceFirst(String.valueOf(str));
						run.setText(runText, 0);
					}
					// 替换1对多
					else if (params.get(matcher.group(1)) instanceof List)
					{
						// 替换1对多文字变量
						if (((List) params.get(matcher.group(1))).get(0) instanceof String)
						{
							List<String> list = (List) params.get(matcher.group(1));
							boolean first = true;
							int index = 1;
							for (String str : list)
							{
								if (first)
								{
									runText = index++ + "、" + matcher.replaceFirst(String.valueOf(str));
									run.setText(runText, 0);
									run.addBreak(BreakType.TEXT_WRAPPING);
									first = false;
								}
								else
								{
									run.setText(index++ + "、" + str);
									run.addBreak(BreakType.TEXT_WRAPPING);
								}
							}
						}
						// 替换1对多图片变量
						else if (((List<?>) params.get(matcher.group(1))).get(0) instanceof Map)
						{
							String nameString = (String)((Map)((List) params.get(matcher.group(1))).get(0)).get("contents");
							List<String> contents = (List<String>) params.get(nameString);
							Map<String, Object> picListmap = (Map) ((List) params.get(matcher.group(1))).get(1);
							boolean first = true;
							for (String str : contents)
							{
								if (first)
								{
									runText = matcher.replaceFirst(String.valueOf(str));
									run.setText(runText, 0);
									run.addBreak(BreakType.PAGE);
									List<String> picList = (List<String>) picListmap.get(str);
									for (String imgFile : picList)
									{
										addPic(imgFile, run);
									}
									first = false;
								}
								else
								{
									// 让名字居中
									run.addBreak(BreakType.TEXT_WRAPPING);
									run.addBreak(BreakType.TEXT_WRAPPING);
									run.addBreak(BreakType.TEXT_WRAPPING);
									run.addBreak(BreakType.TEXT_WRAPPING);
									run.addBreak(BreakType.TEXT_WRAPPING);
									run.addBreak(BreakType.TEXT_WRAPPING);
									
									run.setText(str);
									run.addBreak(BreakType.PAGE);

									List<String> picList = (List<String>) picListmap.get(str);
									for (String imgFile : picList)
									{
										addPic(imgFile, run);
									}
								}
							}
						}
					}
				}
			}
		}
	}
	
	/** 替换文本内容 */
	private void replaceWordTextJust(XWPFParagraph para, Map<String, Object> params) throws Exception
	{
		List<XWPFRun> runs;
		Matcher matcher;
		
		if (this.matcher(para.getParagraphText()).find())
		{
			runs = para.getRuns();
			for (int i = 0; i < runs.size(); i++)
			{
				XWPFRun run = runs.get(i);
				String runText = run.toString();
				matcher = this.matcher(runText);

				if (matcher.find())
				{
					// 替换1对1变量
					if (params.get(matcher.group(1)) instanceof String)
					{
						String str = (String) params.get(matcher.group(1));
						runText = matcher.replaceFirst(String.valueOf(str));
						run.setText(runText, 0);
					}
				}
			}
		}
	}
	
	/** 替换表格中文本内容 */
	@SuppressWarnings("unchecked")
	private void replaceTableWordTextJust(XWPFTable table, Map<String, Object> params) throws Exception
	{
		Matcher matcher;
		List<Map<String, String>> hasCellContant = new ArrayList<Map<String, String>>();
		boolean isList = false;
		boolean noSN = false;
		
		int count = table.getNumberOfRows();
		
		String textString = table.getText();
		if (this.matcher(textString).find())
		{
			for (int i = 0; i < count; i++)
			{
				XWPFTableRow row = table.getRow(i);
				List<XWPFTableCell> cells = row.getTableCells();
				for (int j = 0; j < cells.size(); j++)
				{
					// 某个单元格
					XWPFTableCell cell = cells.get(j);
					// 单元格内容
					String cellTextString = cell.getText();
					matcher = this.matcher(cellTextString);
					
					if (matcher.find())
					{
						// 标签名字
						String paramName = EL_LEFT + matcher.group(1) + EL_RIGHT;
						String replaceText = "";
						// 替换1对1变量
						if (params.get(matcher.group(1)) instanceof String)
						{
							replaceText = (String) params.get(matcher.group(1));
						}
						else if (params.get(matcher.group(1)) instanceof List) 
						{
//							if (getNoSNTable().contains(entry.getKey()))
//								noSN = true;
//							}
							List<Map<String, String>> list = (List<Map<String, String>>) params.get(matcher.group(1));
							hasCellContant.addAll(list);
							isList = true;
						}
						replaceText(cell, paramName, replaceText);
						if (isList) break;
					}
				}
				if (isList) break;
			}
			if (isList) addListValueInTable(hasCellContant, table, noSN);
		}
	}
	
	/**
	 * Read Word Table Range Rule
	 * 
	 * @param table
	 * @param mapProperties
	 */
	@SuppressWarnings("unchecked")
	private void readWordTableRangeRule(XWPFTable table, Map<String, Object> mapProperties)
	{
		List<Map<String, String>> hasCellContant = new ArrayList<Map<String, String>>();
		int rcount = table.getNumberOfRows();
		boolean isListValue = false;
		boolean noSN = false;
		for (int i = 0; i < rcount; i++)
		{
			// 行
			XWPFTableRow row = table.getRow(i);
			// 单元格列表
			List<XWPFTableCell> cells = row.getTableCells();
			for (int j = 0; j < cells.size(); j++)
			{
				// 某个单元格
				XWPFTableCell cell = cells.get(j);
				// 单元格内容
				String cellTextString = cell.getText();
				for (Map.Entry<String, Object> entry : mapProperties.entrySet())
				{
					// 标签名字
					String paramName = EL_LEFT + entry.getKey() + EL_RIGHT;
					// 如果单元格内容里面包含了标签名
					if (cellTextString.contains(paramName))
					{
						String replaceText = "";
						if (entry.getValue() instanceof List)
						{
							if (getNoSNTable().contains(entry.getKey()))
							{
								noSN = true;
							}
							List<Map<String, String>> list = (List<Map<String, String>>) entry.getValue();
							hasCellContant.addAll(list);
							isListValue = true;
						}
						else
						{
							if (null != entry.getValue())
							{
								replaceText = String.valueOf(entry.getValue());
							}
						}
						replaceText(cell, paramName, replaceText);
						if (isListValue) break;
					}
				}
				if (isListValue) break;
			}
			if (isListValue) break;
		}
		if (isListValue) addListValueInTable(hasCellContant, table, noSN);
	}
	
	private void addListValueInTable(List<Map<String, String>> hasCellContant, XWPFTable table, boolean noSN)
	{
		Matcher matcher;
		
		// add text of list values in table
		Map<Integer, List<String>> insertTextIndex = new HashMap<Integer, List<String>>();
		if (hasCellContant.size() > 0)
		{
			// 从表头的下一行开始取，直到大于this.maxTableRowCount开始创建
			int rowIndex = 1;
			for (int i = 0; i < hasCellContant.size(); i++)
			{
				XWPFTableRow row = null;
				List<XWPFTableCell> cells = null;
				// 表格替换内容
				Map<String, String> map = hasCellContant.get(i);
				
				if (rowIndex <= this.maxTableRowCount)
				{
					// 第一次必须取到序号“1”用于替换
					if (rowIndex == 1)
					{
						String ttString = "";
						do
						{
							row = table.getRow(rowIndex++);
							ttString = row.getTableCells().get(0).getText();
						}
						while (ttString.equals(""));
					}
					else {
						row = table.getRow(rowIndex++);
					}
					
				}
				else
				{
					row = table.createRow();
				}
				cells = row.getTableCells();
				// 表格中的一行进行替换
				for (int j = 0; j < cells.size(); j++)
				{
					// 每个单元格
					XWPFTableCell cell = cells.get(j);
					String text = cell.getText();
					// 第一列
					if (j == 0 && i > 0)
					{
						if (!noSN)
						{
							text = String.valueOf(i + 1);
						}
					}
					// 其他列
					else if (j != 0)
					{
						List<String> listParamNames = new ArrayList<String>();
						
						// First line
						if (i == 0)
						{
							matcher = this.matcher(text);
							if (matcher.find() && map.get(matcher.group(1)) != null)
							{
								String paramName = EL_LEFT + matcher.group(1) + EL_RIGHT;
								text = text.replace(paramName, map.get(matcher.group(1)));
								listParamNames.add(matcher.group(1));
								insertTextIndex.put(j, listParamNames);
							}
						}
						// Other line
						else
						{
							listParamNames = insertTextIndex.get(j);
							if (listParamNames != null)
								for (String name : listParamNames)
								{
									text += map.get(name);
								}
						}
					}

					XWPFParagraph cellParagraph = cell.getParagraphs().get(0);
					XWPFRun cellRun;
					
					if (i == 0)
					{
						if (j != 0)
						{
							cellRun = cellParagraph.getRuns().get(0);
							cellRun.setFontSize(11);
							cellRun.setText(text, 0);							
						}
					}
					else
					{
						cellRun = cellParagraph.createRun();
						cellRun.setFontSize(11);
						cellRun.setFontFamily("仿宋");
						cellRun.setText(text, 0);
					}
				}
			}
		}
	}

	
	/** 向word文档中追加企业人员册内容 */
	@SuppressWarnings("unchecked")
	public void appendStaffContent(Map<String, Object> params) throws Exception
	{
		Map<String, String> argsMap = new HashMap<String, String>();
		argsMap.put("constructionStaffLevel1Name", "一级建造师");
		argsMap.put("constructionStaffLevel2Name", "二级建造师");
		argsMap.put("technicalStaffTitle9Name", "高级工程师");
		argsMap.put("technicalStaffTitle10Name", "工程师");
		
		// 按照此顺序来追加内容
		List<String> templatelist = new ArrayList<String>();
		templatelist.add("constructionStaffLevel1Name");
		templatelist.add("constructionStaffLevel1Photo");
		templatelist.add("constructionStaffLevel2Name");
		templatelist.add("constructionStaffLevel2Photo");
		templatelist.add("technicalStaffTitle9Name");
		templatelist.add("technicalStaffTitle9Photo");
		templatelist.add("technicalStaffTitle10Name");
		templatelist.add("technicalStaffTitle10Photo");
		
		for (String template : templatelist)
		{
			if (params.get(template) instanceof List && ((List<?>)params.get(template)).size() > 0)
			{
				// 替换1对多文字变量
				if (((List<?>) params.get(template)).get(0) instanceof String)
				{
					List<String> list = (List<String>) params.get(template);
					
					XWPFParagraph p = document.createParagraph();
					XWPFRun r = p.createRun();
					
					r.setFontSize(26);
					r.setText(argsMap.get(template));
					r.setFontFamily("宋体");

					p = document.createParagraph();
					r = p.createRun();
					r.setFontSize(24);
					
					int index = 1;
					for (String string : list)
					{
						r.setText(index++ + "、" + string);
						r.addCarriageReturn();
					}
					r.addBreak(BreakType.PAGE);
				}
				// 替换1对多图片变量
				else if (((List<?>) params.get(template)).get(0) instanceof Map)
				{
					String nameString = (String)((Map<?, ?>)((List<?>) params.get(template)).get(0)).get("contents");
					List<String> contents = (List<String>) params.get(nameString);
					Map<String, Object> picListmap = (Map<String, Object>) ((List<?>) params.get(template)).get(1);
					
					for (String str : contents)
					{
						XWPFParagraph p = document.createParagraph();
						p.setAlignment(ParagraphAlignment.CENTER);
						XWPFRun r = p.createRun();
						r.setFontSize(26);
						
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.setText(str);
						r.addBreak(BreakType.PAGE);
						
						List<String> picList = (List<String>) picListmap.get(str);
						for (String imgFile : picList)
						{
							addPic(imgFile, r);
						}
					}
				}
			}
		}
	}

	/** 向word文档中追加企业综合册内容 */
	@SuppressWarnings("unchecked")
	public void appendBasicContent(Map<String, Object> params) throws Exception
	{
		// 按照此顺序来追加内容
		List<String> templatelist = new ArrayList<String>();
		templatelist.add("basicInfoName");
		templatelist.add("basicInfoPhoto");
		
		for (String template : templatelist)
		{
			if (params.get(template) instanceof List && ((List<?>)params.get(template)).size() > 0)
			{
				// 替换1对多文字变量
				if (((List<?>) params.get(template)).get(0) instanceof String)
				{
					List<String> list = (List<String>) params.get(template);
					
					XWPFParagraph p = document.createParagraph();
					XWPFRun r = p.createRun();
					
					r.setFontSize(20);
					
					int index = 1;
					for (String string : list)
					{
						r.setText(index++ + "、" + string);
						r.addCarriageReturn();
					}
					r.addBreak(BreakType.PAGE);
				}
				// 替换1对多图片变量
				else if (((List<?>) params.get(template)).get(0) instanceof Map)
				{
					String nameString = (String)((Map<?, ?>)((List<?>) params.get(template)).get(0)).get("contents");
					List<String> contents = (List<String>) params.get(nameString);
					Map<String, Object> picListmap = (Map<String, Object>) ((List<?>) params.get(template)).get(1);
					
					for (String str : contents)
					{
						XWPFParagraph p = document.createParagraph();
						p.setAlignment(ParagraphAlignment.CENTER);
						XWPFRun r = p.createRun();
						r.setFontSize(26);
						
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.addCarriageReturn();
						r.setText(str);
						r.addBreak(BreakType.PAGE);
						
						List<String> picList = (List<String>) picListmap.get(str);
						for (String imgFile : picList)
						{
							addPic(imgFile, r);
						}
					}
				}
			}
		}
	}
	
	/** 生成word文档 */
	public void writeDoc(String outputPath) throws Exception
	{
		File outputFile =  new File(outputPath);
		OutputStream os = new FileOutputStream(outputFile);
		document.write(os);
		os.flush();
		this.close(os);
		this.close(is);
	}
	
	/** 生成word文档 */
	public void writeDoc(HttpServletResponse response, String fileName) throws Exception
	{
		// 设置Response参数
		response.setContentType("application/vnd.ms-word");
		response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes("GBK"), "ISO8859-1"));
		
		OutputStream os = response.getOutputStream();
		document.write(os);
		os.flush();
		this.close(os);
		this.close(is);
	}
	
	private void addPic(String imgFile, XWPFRun r) throws Exception
	{
		ImageInfo imageInfo = getImageInfo(imgFile);

		r.addPicture(new FileInputStream(imgFile), imageInfo.getType(), imgFile, Units.toEMU(imageInfo.getWidth()), Units.toEMU(imageInfo.getHeight())); 
		r.addBreak(BreakType.PAGE);
	}

	/**
	 * 根据图片类型，取得对应的图片类型代码
	 * 
	 * @param picType
	 * @return int
	 */
	private int getPictureType(String picType)
	{
		int res = XWPFDocument.PICTURE_TYPE_PICT;
		if (picType != null)
		{
			if (picType.endsWith("png"))
			{
				res = XWPFDocument.PICTURE_TYPE_PNG;
			}
			else if (picType.endsWith("jpg") || picType.endsWith("jpeg"))
			{
				res = XWPFDocument.PICTURE_TYPE_JPEG;
			}
			else if (picType.endsWith("gif"))
			{
				res = XWPFDocument.PICTURE_TYPE_GIF;
			}
		}
		return res;
	}

	public ImageInfo getImageInfo(String path) throws Exception
	{
		File imgfile = new File(path);

		// 图片缩放比例
		double ratio = 0.0;

		BufferedImage Bi = ImageIO.read(imgfile);

		double width = Bi.getWidth();
		double height = Bi.getHeight();

		// 高度超出
		if ((height > WordUtil.HEIGHT_MAX_SIZE))
		{
			ratio = WordUtil.HEIGHT_MAX_SIZE / height;
			height = WordUtil.HEIGHT_MAX_SIZE;
			width = ratio * width;
		}
		
		if ((width > WordUtil.WIDTH_MAX_SIZE))
		{
			ratio = WordUtil.WIDTH_MAX_SIZE / width;
			height = ratio * height;
			width = WordUtil.WIDTH_MAX_SIZE;
		}

		ImageInfo imageInfo = new ImageInfo();
		imageInfo.setType(getPictureType(path));
		imageInfo.setWidth((int) width);
		imageInfo.setHeight((int) height);

		return imageInfo;
	}
	
	/**
	 * 正则匹配字符串
	 * 
	 * @param str
	 * @return
	 */
	private Matcher matcher(String str)
	{
		Pattern pattern = Pattern.compile("\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
	/**
	 * 关闭输入流
	 * 
	 * @param is
	 */
	private void close(InputStream is)
	{
		if (is != null)
		{
			try
			{
				is.close();
			}
			catch (IOException e)
			{
				e.printStackTrace();
			}
		}
	}

	/**
	 * 关闭输出流
	 * 
	 * @param os
	 */
	private void close(OutputStream os)
	{
		if (os != null)
		{
			try
			{
				os.close();
			}
			catch (IOException e)
			{
				e.printStackTrace();
			}
		}
	}
	
	@SuppressWarnings("deprecation")
	private void replaceText(XWPFTableCell cell, String paramName, String replaceText)
	{
		if (null != cell.getParagraphs())
		{
			int lengthOld = cell.getParagraphs().size();
			for (int i = 0; i < lengthOld; i++)
			{
				XWPFParagraph xWPFParagraph = cell.getParagraphs().get(i);
				for (int j = 0; j < xWPFParagraph.getCTP().getRArray().length; j++)
				{
					CTR run = xWPFParagraph.getCTP().getRArray()[j];
					for (int k = 0; k < run.getTArray().length; k++)
					{
						CTText text = run.getTArray()[k];
						String stringValue = text.getStringValue();
						if (stringValue.contains(paramName))
						{
							stringValue = stringValue.replace(paramName, replaceText);
							text.setStringValue(stringValue);
						}
					}
				}
			}
		}
	}

	public void setNoSNTable(List<String> noSNTable)
	{
		this.noSNTable = noSNTable;
	}

	public List<String> getNoSNTable()
	{
		if (null == noSNTable)
		{
			noSNTable = new ArrayList<String>();
		}
		return noSNTable;
	}
	
	public int getMaxTableRowCount()
	{
		return maxTableRowCount;
	}

	public void setMaxTableRowCount(int maxTableRowCount)
	{
		this.maxTableRowCount = maxTableRowCount;
	}
}

class ImageInfo
{
	private int width;
	private int height;
	private int type;

	public int getWidth()
	{
		return width;
	}

	public void setWidth(int width)
	{
		this.width = width;
	}

	public int getHeight()
	{
		return height;
	}

	public void setHeight(int height)
	{
		this.height = height;
	}

	public int getType()
	{
		return type;
	}

	public void setType(int type)
	{
		this.type = type;
	}

}
