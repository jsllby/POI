package com.ruite.util.poi;

import java.awt.image.BufferedImage;
import java.io.FileOutputStream;

import org.jbarcode.JBarcode;
import org.jbarcode.encode.Code128Encoder;
import org.jbarcode.encode.InvalidAtributeException;
import org.jbarcode.paint.BaseLineTextPainter;
import org.jbarcode.paint.WidthCodedPainter;
import org.jbarcode.util.ImageUtil;

/**
 * 条形码生成器
 * @author fangzhiyang  
 *
 */
public class BarcodeUtil {

	private JBarcode localJBarcode;
	private BufferedImage localBufferedImage = null;

	public BarcodeUtil() {
		localJBarcode = new JBarcode(Code128Encoder.getInstance(), WidthCodedPainter.getInstance(),
				BaseLineTextPainter.getInstance());
	}

	// private void saveToJPEG(BufferedImage paramBufferedImage, String
	// paramString) {
	// saveToFile(paramBufferedImage, paramString, "jpeg");
	// }

	public String saveToPNG(String code, String paramString) {
//		paramString = paramString.replace("/", "");
//		paramString = paramString.replace("\\", "");
		
		try {
			localBufferedImage = localJBarcode.createBarcode(code);
		} catch (InvalidAtributeException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		saveToFile(localBufferedImage, paramString, "png");
		return paramString;
	}

	// private void saveToGIF(BufferedImage paramBufferedImage, String
	// paramString) {
	// saveToFile(paramBufferedImage, paramString, "gif");
	// }

	private void saveToFile(BufferedImage paramBufferedImage, String paramString1, String paramString2) {
		try {
			FileOutputStream localFileOutputStream = new FileOutputStream(paramString1);
			ImageUtil.encodeAndWrite(paramBufferedImage, paramString2, localFileOutputStream, 96, 96);
			localFileOutputStream.close();
		} catch (Exception localException) {
			localException.printStackTrace();
		}
	}

}
