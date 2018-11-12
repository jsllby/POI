package com.ruite.util.poi;

/**
 * 暂付款报销联
 * @author fangzhiyang
 *
 */
public class LocationMap1 extends LocationMap
{
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
	public static final String COLLEGE_2 = "单位2";
	public static final String NAME_2 = "姓名2";
	public static final String CCB_ACCOUNT_2 = "建行2";
	public static final String ICB_ACCOUNT_2 = "工行2";
	public static final String JOBNUM_2 = "工号2";
	public static final String TELE_2 = "电话2";
	public static final String ITEMCODE_2 = "项目代码2";
	public static final String ITEMNAME_2 = "项目名称2";
	public static final String REASON_2 = "事由2";
	public static final String AMOUNT_2 = "金额2";
	public static final String WAN_2 = "万2";
	public static final String QIAN_2 = "千2";
	public static final String BAI_2 = "百2";
	public static final String SHI_2 = "十2";
	public static final String YUAN_2 = "元2";
	public static final String JIAO_2 = "角2";
	public static final String FEN_2 = "分2";
	public static final String AGENT_2 = "经办人2";
	public static final String RECEIPT_NUM_2 = "发票张数2";
	public static final String DATE_2 = "日期2";
	
	public LocationMap1() {
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
		locationMap.put(WAN, new MyData(7, 11));
		locationMap.put(QIAN, new MyData(7, 12));
		locationMap.put(BAI, new MyData(7, 13));
		locationMap.put(SHI, new MyData(7, 14));
		locationMap.put(YUAN, new MyData(7, 15));
		locationMap.put(JIAO, new MyData(7, 16));
		locationMap.put(FEN, new MyData(7, 17));

		// 第四行
		locationMap.put(PAYEE, new MyData(11, 3));
		locationMap.put(BANK, new MyData(11, 6));
		locationMap.put(ACCOUNT, new MyData(11, 9));
		locationMap.put(AGENT, new MyData(19, 3));
		locationMap.put(RECEIPT_NUM, new MyData(1, 15));
		locationMap.put(DATE, new MyData(1, 3));

		//冲账
		// 第一行
		locationMap.put(COLLEGE_2, new MyData(29, 1));
		locationMap.put(NAME_2, new MyData(29, 3));
		locationMap.put(CCB_ACCOUNT_2, new MyData(29, 6));
		locationMap.put(ICB_ACCOUNT_2, new MyData(30, 6));
		locationMap.put(REASON_2, new MyData(29, 10));

		// 第二行
		locationMap.put(JOBNUM_2, new MyData(31, 1));
		locationMap.put(TELE_2, new MyData(31, 3));
		locationMap.put(ITEMCODE_2, new MyData(31, 6));
		locationMap.put(ITEMNAME_2, new MyData(32, 6));

		// 第三行
		locationMap.put(AMOUNT_2, new MyData(33, 4));
		locationMap.put(WAN_2, new MyData(34, 11));
		locationMap.put(QIAN_2, new MyData(34, 12));
		locationMap.put(BAI_2, new MyData(34, 13));
		locationMap.put(SHI_2, new MyData(34, 14));
		locationMap.put(YUAN_2, new MyData(34, 15));
		locationMap.put(JIAO_2, new MyData(34, 16));
		locationMap.put(FEN_2, new MyData(34, 17));

		// 第四行
		locationMap.put(AGENT_2, new MyData(45, 3));
		locationMap.put(RECEIPT_NUM_2, new MyData(28, 15));
		locationMap.put(DATE_2, new MyData(28, 3));
	}
}
