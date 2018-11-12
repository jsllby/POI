package com.ruite.util.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import com.ruite.UserException;
import com.ruite.file.FileManager;

public class ToScufdExcel {

	public static int dpi = 150; // 条码参数
	private String workNo;

	public Workbook workbook;
	public LocationMap locationMap;

	public static void main(String[] args) throws Exception {

	}

	public ToScufdExcel(Map<String, String> textMap, String wn, int panel) throws UserException {
		if (wn == null || wn.equals(""))
			wn = "default";

		workNo = wn;

		switch (panel) {
		case 0:
			locationMap = new LocationMap0();
			break;
		case 1:
			locationMap = new LocationMap1();
			break;
		case 2:
			locationMap = new LocationMap2();
			break;
		case 3:
			locationMap = new LocationMap3();
			break;
		case 4:
			locationMap = new LocationMap4();
			break;
		case 5:
			locationMap = new LocationMap5();
			break;
		default:
			throw new UserException("参数有误");
		}

		try {
			workbook = WorkbookFactory
					.create(new File(FileManager.getInstance().getTemplateHome() + File.separator + (panel + 1) + ".xls"));
			Sheet sheet = workbook.getSheetAt(0);

			// 给参数列表赋值
			List<MyData> params = new ArrayList<MyData>(locationMap.getValues());
			
			for (String key : textMap.keySet()) {
				if (!key.equals(LocationMap.REAL_AMOUNT)) {
					MyData data = locationMap.getData(key);
					data.setText(textMap.get(key));
				}
			}

			// 存放条码图片的路径
			String barcodePicPath = FileManager.getInstance().getTemplateHome() + File.separator + workNo;
			File rootPath = new File(barcodePicPath);
			if (!rootPath.exists())
				rootPath.mkdirs();

			// 生成条码
			String bcVal = textMap.get(LocationMap.ITEMCODE);
			String bcVal2 = textMap.get(LocationMap.REAL_AMOUNT);
			// 生成条码图片
			BarcodeUtil util = new BarcodeUtil();
			String bcValWithoutLine = bcVal.replace("/", "").replace("\\", "");
			String bcValWithoutLine2 = bcVal2.replace("/", "").replace("\\", "");
			
			String bcPath = util.saveToPNG(bcVal, barcodePicPath + File.separator + bcValWithoutLine + ".png");
			String bcPath2 = util.saveToPNG(bcVal2, barcodePicPath + File.separator + bcValWithoutLine2 + ".png");
			
			// 插入条码
			insertBarcode(workbook, sheet, bcPath, bcPath2, panel);
			// 冲账需要再次插入条码
			if (panel == 1 || panel == 2)
				insertBarcode1(workbook, sheet, bcPath, bcPath2, panel);
				
			// 插入文本
			insertText(sheet, params);
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/** 输出excel */
	public void writeExcel(HttpServletResponse response, String fileName) throws IOException {
		// 设置Response参数
		response.setContentType("application/vnd.ms-excel");
		response.setHeader("Content-Disposition",
				"attachment;filename=" + new String(fileName.getBytes("gb2312"), "ISO8859-1"));

		if (this.workbook != null) {
			OutputStream os = response.getOutputStream();
			try {
				this.workbook.write(os);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				if (os != null) {
					os.close();
				}
			}
		}
	}

	/** 向excel当中插入条码 */
	private void insertBarcode(Workbook wb, Sheet st, String barcodePic, String barcodePic2, int panel) {
		InputStream inputStream = null;
		InputStream inputStream2 = null;

		try {
			inputStream = new FileInputStream(barcodePic);
			byte[] bytes = IOUtils.toByteArray(inputStream);
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

			inputStream2 = new FileInputStream(barcodePic2);
			byte[] bytes2 = IOUtils.toByteArray(inputStream2);
			int pictureIdx2 = wb.addPicture(bytes2, Workbook.PICTURE_TYPE_PNG);

			CreationHelper helper = wb.getCreationHelper();
			Drawing drawing = st.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();

			// 判断sheet，插入条形码
			if (panel == 0) {
				anchor.setCol1(2);
				anchor.setRow1(21);
			} else if (panel == 1) {
				anchor.setCol1(3);
				anchor.setRow1(21);
			} else if (panel == 2) {
				anchor.setCol1(2);
				anchor.setRow1(23);
			} else if (panel == 3) {
				anchor.setCol1(2);
				anchor.setRow1(22);
			} else if (panel == 4) {
				anchor.setCol1(4);
				anchor.setRow1(25);
			} else {
				anchor.setCol1(1);
				anchor.setRow1(19);
			}

			Picture pict = drawing.createPicture(anchor, pictureIdx);
			pict.resize();

			ClientAnchor anchor2 = helper.createClientAnchor();
			if (panel == 0) {
				anchor2.setCol1(6);
				anchor2.setRow1(21);
			} else if (panel == 1) {
				anchor2.setCol1(8);
				anchor2.setRow1(21);
			} else if (panel == 2) {
				anchor2.setCol1(7);
				anchor2.setRow1(23);
			} else if (panel == 3) {
				anchor2.setCol1(7);
				anchor2.setRow1(22);
			} else if (panel == 4) {
				anchor2.setCol1(8);
				anchor2.setRow1(25);
			} else {
				anchor2.setCol1(6);
				anchor2.setRow1(19);
			}

			pict = drawing.createPicture(anchor2, pictureIdx2);
			pict.resize();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (inputStream != null) {
					inputStream.close();
				}
				if (inputStream2 != null) {
					inputStream2.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	/** 向excel中再次插入条码 */
	private void insertBarcode1(Workbook wb, Sheet st, String barcodePic, String barcodePic2, int i) {
		InputStream inputStream = null;
		InputStream inputStream2 = null;

		try {
			inputStream = new FileInputStream(barcodePic);
			byte[] bytes = IOUtils.toByteArray(inputStream);
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

			inputStream2 = new FileInputStream(barcodePic2);
			byte[] bytes2 = IOUtils.toByteArray(inputStream2);
			int pictureIdx2 = wb.addPicture(bytes2, Workbook.PICTURE_TYPE_PNG);

			CreationHelper helper = wb.getCreationHelper();
			Drawing drawing = st.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();

			//判断sheet，插入条形码
			if (i == 1) {
				anchor.setCol1(2);
				anchor.setRow1(46);
			} else if (i == 2) {
				anchor.setCol1(2);
				anchor.setRow1(51);
			}
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			pict.resize();
			ClientAnchor anchor2 = helper.createClientAnchor();
			if (i == 1) {
				anchor2.setCol1(8);
				anchor2.setRow1(46);
			} else if (i == 2) {
				anchor2.setCol1(8);
				anchor2.setRow1(51);
			}
			pict = drawing.createPicture(anchor2, pictureIdx2);
			pict.resize();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (inputStream != null) {
					inputStream.close();
				}
				if (inputStream2 != null) {
					inputStream2.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/** 通过使用模板生成Excel文件,模板当中包含样式, 这样我们只为模板填充数据就可以有相应的样式 */
	private void insertText(Sheet sheet, List<MyData> params) {
		// 分别替换内容
		for (MyData myData : params) {
			Row row = sheet.getRow(myData.getRow());
			Cell cell = row.getCell(myData.getCell());
			cell.setCellValue(myData.getText());
		}
	}

	/** 金额转大写 */
	public static String digitUppercase(String num) throws Exception {
		if (num == null || num.equals(""))
			return "";

		String digit[] = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
		/**
		 * 仟 佰 拾 ' ' ' ' $4 $3 $2 $1 万 $8 $7 $6 $5 亿 $12 $11 $10 $9
		 */
		String unit1[] = { "", "拾", "佰", "仟" };// 把钱数分成段,每四个一段,实际上得到的是一个二维数组
		String unit2[] = { "元", "万", "亿", "万亿" }; // 把钱数分成段,每四个一段,实际上得到的是一个二维数组
		BigDecimal bigDecimal = new BigDecimal(num);
		bigDecimal = bigDecimal.multiply(new BigDecimal(100));
		// Double bigDecimal = new Double(name*100); 存在精度问题 eg：145296.8
		String strVal = String.valueOf(bigDecimal.toBigInteger());
		String head = strVal.substring(0, strVal.length() - 2); // 整数部分
		String end = strVal.substring(strVal.length() - 2); // 小数部分
		String endMoney = "";
		String headMoney = "";
		if ("00".equals(end)) {
			endMoney = "整";
		} else {
			if (!end.substring(0, 1).equals("0")) {
				endMoney += digit[Integer.valueOf(end.substring(0, 1))] + "角";
			} else if (end.substring(0, 1).equals("0") && !end.substring(1, 2).equals("0")) {
				endMoney += "零";
			}
			if (!end.substring(1, 2).equals("0")) {
				endMoney += digit[Integer.valueOf(end.substring(1, 2))] + "分";
			}
		}
		char[] chars = head.toCharArray();
		Map<String, Boolean> map = new HashMap<String, Boolean>();// 段位置是否已出现zero
		boolean zeroKeepFlag = false;// 0连续出现标志
		int vidxtemp = 0;
		for (int i = 0; i < chars.length; i++) {
			int idx = (chars.length - 1 - i) % 4;// 段内位置 unit1
			int vidx = (chars.length - 1 - i) / 4;// 段位置 unit2
			String s = digit[Integer.valueOf(String.valueOf(chars[i]))];
			if (!"零".equals(s)) {
				headMoney += s + unit1[idx] + unit2[vidx];
				zeroKeepFlag = false;
			} else if (i == chars.length - 1 || map.get("zero" + vidx) != null) {
				headMoney += "";
			} else {
				headMoney += s;
				zeroKeepFlag = true;
				map.put("zero" + vidx, true);// 该段位已经出现0；
			}
			if (vidxtemp != vidx || i == chars.length - 1) {
				headMoney = headMoney.replaceAll(unit2[vidx], "");
				headMoney += unit2[vidx];
			}
			if (zeroKeepFlag && (chars.length - 1 - i) % 4 == 0) {
				headMoney = headMoney.replaceAll("零", "");
			}
		}
		return headMoney + endMoney;
	}

	/** 生成7位金额 */
	public static String fillTextMapByDigit(String num) {
		if (num == null || num.equals(""))
			return "0000000";
		BigDecimal bigDecimal = new BigDecimal(num);
		bigDecimal = bigDecimal.multiply(new BigDecimal(100));
		String strVal = String.format("%07d", bigDecimal.toBigInteger());
		// String head = strVal.substring(0, strVal.length() - 2); // 整数部分
		// String end = strVal.substring(strVal.length() - 2); // 小数部分
		return strVal;
	}
	
	/** 生成11位金额 */
	public static String fillTextMapByElevenDigit(String num) {
		if (num == null || num.equals(""))
			return "00000000000";
		BigDecimal bigDecimal = new BigDecimal(num);
		bigDecimal = bigDecimal.multiply(new BigDecimal(100));
		String strVal = String.format("%11d", bigDecimal.toBigInteger());
		return strVal;
	}

}