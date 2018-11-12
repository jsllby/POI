package com.ruite.util.poi;

/**
 * 大额款项支(借)款申请单
 * @author fangzhiyang
 *
 */
public class LocationMap3 extends LocationMap {
	public static final String COLLEGE = "单位";
	public static final String NAME = "姓名";
	public static final String CCB_ACCOUNT = "建行";
	public static final String ICB_ACCOUNT = "工行";
	public static final String JOBNUM = "工号";
	public static final String TELE = "电话";
	public static final String ITEMCODE = "项目代码";
	public static final String ITEMNAME = "项目名称";
	public static final String REASON = "事由";
	public static final String AMOUNT = "金额";
	public static final String YI = "亿";
	public static final String QIANWAN = "千万";
	public static final String BAIWAN = "百万";
	public static final String SHIWAN = "十万";
	public static final String WAN = "万";
	public static final String QIAN = "千";
	public static final String BAI = "百";
	public static final String SHI = "十";
	public static final String YUAN = "元";
	public static final String JIAO = "角";
	public static final String FEN = "分";
	public static final String PAYEE = "收款单位";
	public static final String BANK = "银行";
	public static final String ACCOUNT = "账号";
	public static final String AGENT = "经办人";
	public static final String RECEIPT_NUM = "发票张数";

	public LocationMap3() {
		/** 大额支借款申请单 */
		// 第一行
		locationMap.put(COLLEGE, new MyData(2, 1));
		locationMap.put(NAME, new MyData(2, 3));
		locationMap.put(CCB_ACCOUNT, new MyData(2, 6));
		locationMap.put(ICB_ACCOUNT, new MyData(3, 6));
		locationMap.put(REASON, new MyData(2, 10));

		// 第二行
		locationMap.put(JOBNUM, new MyData(4, 1));
		locationMap.put(TELE, new MyData(4, 3));
		locationMap.put(ITEMCODE, new MyData(4, 6));
		locationMap.put(ITEMNAME, new MyData(5, 6));

		// 第三行
		locationMap.put(AMOUNT, new MyData(6, 4));
		locationMap.put(YI, new MyData(7, 9));
		locationMap.put(QIANWAN, new MyData(7, 10));
		locationMap.put(BAIWAN, new MyData(7, 11));
		locationMap.put(SHIWAN, new MyData(7, 12));
		locationMap.put(WAN, new MyData(7, 13));
		locationMap.put(QIAN, new MyData(7, 14));
		locationMap.put(BAI, new MyData(7, 15));
		locationMap.put(SHI, new MyData(7, 16));
		locationMap.put(YUAN, new MyData(7, 17));
		locationMap.put(JIAO, new MyData(7, 18));
		locationMap.put(FEN, new MyData(7, 19));

		// 第四行
		locationMap.put(PAYEE, new MyData(11, 3));
		locationMap.put(BANK, new MyData(11, 6));
		locationMap.put(ACCOUNT, new MyData(11, 9));
		locationMap.put(AGENT, new MyData(21, 2));
		locationMap.put(RECEIPT_NUM, new MyData(1, 16));
		locationMap.put(DATE, new MyData(1, 3));
	}

}
