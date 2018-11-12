package com.ruite.util.poi;

/**
 * 内部转账单
 * @author fangzhiyang
 *
 */
public class LocationMap5 extends LocationMap
{
	public static final String PAYER = "付款单位";
	public static final String PAYEE = "收款单位";	
	public static final String TOPIC_1 = "课题名称1";
	public static final String ITEMCODE_1 = "项目代码1";
	public static final String DEPARTMENT_1 = "部门代码1";
	public static final String TOPIC_2 = "课题名称2";
	public static final String ITEMCODE_2 = "项目代码2";
	public static final String DEPARTMENT_2 = "部门代码2";
	public static final String REASON = "事由";
	public static final String SIZE = "规格";
    public static final String AMOUNT_IN_SMALLWORDS = "小写金额";
	public static final String AMOUNT_IN_LARGEWORDS = "大写金额";
	public static final String PAYER_TRANSACTOR = "付款单位经办人";
	public static final String PAYEE_TRANSACTOR = "收款单位经办人";
	public static final String TRANSACTOR_JOBNUM_1 = "经办人工作证号1";
	public static final String TRANSACTOR_JOBNUM_2 = "经办人工作证号2";
	public static final String TELE_1 = "联系电话1";
	public static final String TELE_2 = "联系电话2";
	
    public LocationMap5()
	{
		/** 支借款申请单 */
		locationMap.put(DATE, new MyData(1, 4));
		// 第一行
		locationMap.put(PAYER, new MyData(2, 1));
		locationMap.put(PAYEE, new MyData(2, 7));
		// 第二行
		locationMap.put(TOPIC_1, new MyData(4, 1));
		locationMap.put(ITEMCODE_1, new MyData(5, 5));
		locationMap.put(DEPARTMENT_1, new MyData(4, 5));
		locationMap.put(TOPIC_2, new MyData(4,7));
		locationMap.put(DEPARTMENT_2, new MyData(4,10));
		locationMap.put(ITEMCODE_2, new MyData(5,10));

		// 第三行
		locationMap.put(REASON, new MyData(6, 1));
		
		// 第四行
		locationMap.put(SIZE, new MyData(8, 1));

		// 第五行
		locationMap.put(AMOUNT_IN_LARGEWORDS, new MyData(10, 1));
		locationMap.put(AMOUNT_IN_SMALLWORDS, new MyData(10, 7));
		
		// 第六行
		locationMap.put(PAYER_TRANSACTOR, new MyData(12, 5));
		locationMap.put(PAYEE_TRANSACTOR, new MyData(12, 8));
		locationMap.put(TRANSACTOR_JOBNUM_1, new MyData(14, 5));
		locationMap.put(TRANSACTOR_JOBNUM_2, new MyData(14, 8));
		locationMap.put(TELE_1, new MyData(16, 5));
		locationMap.put(TELE_2, new MyData(16, 8));
	}
    
}