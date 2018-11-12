package com.ruite.util.poi;

/**
 * 大额款项暂付款报销联
 * @author fangzhiyang
 *
 */
public class LocationMap2 extends LocationMap {
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

	//冲账
	public static final String COLLEGE_1 = "单位1";
	public static final String NAME_1 = "姓名1";
	public static final String CCB_ACCOUNT_1 = "建行1";
	public static final String ICB_ACCOUNT_1 = "工行1";
	public static final String JOBNUM_1 = "工号1";
	public static final String TELE_1 = "电话1";
	public static final String ITEMCODE_1 = "项目代码1";
	public static final String ITEMNAME_1 = "项目名称1";
	public static final String REASON_1 = "事由1";
	public static final String AMOUNT_1 = "金额1";
	public static final String YI_1 = "亿1";
	public static final String QIANWAN_1 = "千万1";
	public static final String BAIWAN_1 = "百万1";
	public static final String SHIWAN_1 = "十万1";
	public static final String WAN_1 = "万1";
	public static final String QIAN_1 = "千1";
	public static final String BAI_1 = "百1";
	public static final String SHI_1 = "十1";
	public static final String YUAN_1 = "元1";
	public static final String JIAO_1 = "角1";
	public static final String FEN_1 = "分1";
	public static final String AGENT_1 = "经办人1";
	public static final String RECEIPT_NUM_1 = "发票张数1";
	public static final String DATE_1 = "日期1";

	public LocationMap2() {
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

		/** 大额支借款申请单冲账 */
		// 第一行
		locationMap.put(COLLEGE_1, new MyData(31, 1));
		locationMap.put(NAME_1, new MyData(31, 3));
		locationMap.put(CCB_ACCOUNT_1, new MyData(31, 6));
		locationMap.put(ICB_ACCOUNT_1, new MyData(32, 6));
		locationMap.put(REASON_1, new MyData(31, 10));

		// 第二行
		locationMap.put(JOBNUM_1, new MyData(33, 1));
		locationMap.put(TELE_1, new MyData(33, 3));
		locationMap.put(ITEMCODE_1, new MyData(33, 6));
		locationMap.put(ITEMNAME_1, new MyData(34, 6));

		// 第三行
		locationMap.put(AMOUNT_1, new MyData(35, 4));
		locationMap.put(YI_1, new MyData(36, 9));
		locationMap.put(QIANWAN_1, new MyData(36, 10));
		locationMap.put(BAIWAN_1, new MyData(36, 11));
		locationMap.put(SHIWAN_1, new MyData(36, 12));
		locationMap.put(WAN_1, new MyData(36, 13));
		locationMap.put(QIAN_1, new MyData(36, 14));
		locationMap.put(BAI_1, new MyData(36, 15));
		locationMap.put(SHI_1, new MyData(36, 16));
		locationMap.put(YUAN_1, new MyData(36, 17));
		locationMap.put(JIAO_1, new MyData(36, 18));
		locationMap.put(FEN_1, new MyData(36, 19));

		// 第四行
		locationMap.put(AGENT_1, new MyData(49, 3));
		locationMap.put(RECEIPT_NUM_1, new MyData(30, 16));
		locationMap.put(DATE_1, new MyData(30, 3));
	}
}