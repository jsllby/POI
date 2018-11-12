package com.ruite.util.poi;

/**
 * 差旅费报销单
 * @author fangzhiyang
 *
 */
public class LocationMap4 extends LocationMap
{
	public static final String APPROVER = "批准单位";
	public static final String REASON = "事由";
	public static final String NAME = "姓名";
	public static final String TITLE = "职务职称";
	public static final String ASSISTANCE ="是否报销补助";
	public static final String BILL_AMOUNT ="单据张数";
	public static final String ITEMCODE = "项目编号";
	public static final String UNIT = "派出单位";
	public static final String ITEMNAME= "项目名称";
	public static final String AMOUNT_IN_LARGEWORDS= "大写金额";
	public static final String JOBNUM = "工号";
	public static final String TELE = "电话";
	public static final String CCB_ACCOUNT = "建行卡号";
	public static final String ICB_ACCOUNT = "工行卡号";
	public static final String NAME_1 = "姓名1";
	public static final String TRAVELLING_EXPENSE_SUM = "车船费飞机票合计";
	public static final String CAR_FARE_SUM = "车费合计";
	public static final String ACCOMODATE_AMOUNT_SUM = "住宿金额合计";
	public static final String MEAL_AMOUNT_SUM = "伙食金额合计";
	public static final String TRAVEL_IN_CITY_AMOUNT_SUM = "市内交通金额合计";
	public static final String OTHER_EXPENSE_SUM = "其他费用合计";
	
	public static String[] depar_month = new String[]{"出发月0","出发月1","出发月2","出发月3","出发月4","出发月5","出发月6","出发月7"};
	public static String[] depar_date = new String[]{"出发日0","出发日1","出发日2","出发日3","出发日4","出发日5","出发日6","出发日7"}; 
	public static String[] arri_month = new String[]{"到达月0","到达月1","到达月2","到达月3","到达月4","到达月5","到达月6","到达月7"}; 
	public static String[] arri_date = new String[]{"到达日0","到达日1","到达日2","到达日3","到达日4","到达日5","到达日6","到达日7"}; 
	public static String[] depar_desti = new String[]{"起止地点0","起止地点1","起止地点2","起止地点3","起止地点4","起止地点5","起止地点6","起止地点7"};
	public static String[] travelling_expense = new String[]{"车船费飞机票0","车船费飞机票1","车船费飞机票2","车船费飞机票3","车船费飞机票4","车船费飞机票5","车船费飞机票6","车船费飞机票7"};
	public static String[] car_fare = new String[]{"车费0","车费1","车费2","车费3","车费4","车费5","车费6","车费7"};
	public static String[] accomodate_number = new String[]{"住宿人数0","住宿人数1","住宿人数2","住宿人数3","住宿人数4","住宿人数5","住宿人数6","住宿人数7"};
	public static String[] accomodate_standard = new String[]{"住宿标准0","住宿标准1","住宿标准2","住宿标准3","住宿标准4","住宿标准5","住宿标准6","住宿标准7"};
	public static String[] accomodate_amount = new String[]{"住宿金额0","住宿金额1","住宿金额2","住宿金额3","住宿金额4","住宿金额5","住宿金额6","住宿金额7"};
	public static String[] meal_number = new String[]{"伙食人数0","伙食人数1","伙食人数2","伙食人数3","伙食人数4","伙食人数5","伙食人数6","伙食人数7"};
	public static String[] meal_amount = new String[]{"伙食金额0","伙食金额1","伙食金额2","伙食金额3","伙食金额4","伙食金额5","伙食金额6","伙食金额7"};
	public static String[] travel_in_city_number = new String[]{"市内交通人数0","市内交通人数1","市内交通人数2","市内交通人数3","市内交通人数4","市内交通人数5","市内交通人数6","市内交通人数7"};
	public static String[] travel_in_city_amount = new String[]{"市内交通金额0","市内交通金额1","市内交通金额2","市内交通金额3","市内交通金额4","市内交通金额5","市内交通金额6","市内交通金额7"};
	public static String[] other_expense = new String[]{"其他费用0","其他费用1","其他费用2","其他费用3","其他费用4","其他费用5","其他费用6","其他费用7"};

	
    public LocationMap4()
	{
		/** 差旅费报销单 */
		// 第一行
		locationMap.put(APPROVER, new MyData(2, 4));
		locationMap.put(REASON, new MyData(2, 10));

		// 第二行
		locationMap.put(NAME, new MyData(3, 5));
		
		// 第三行
		locationMap.put(TITLE, new MyData(4, 5));
		
		// 第四行
		locationMap.put(ASSISTANCE, new MyData(7, 15));
		
		// 第五行
		locationMap.put(BILL_AMOUNT, new MyData(11, 16));
		
		//第六行
		locationMap.put(TRAVELLING_EXPENSE_SUM, new MyData(15, 5));
		locationMap.put(CAR_FARE_SUM, new MyData(15, 6));
		locationMap.put(ACCOMODATE_AMOUNT_SUM, new MyData(15, 9));
		locationMap.put(MEAL_AMOUNT_SUM , new MyData(15, 11));
		locationMap.put(TRAVEL_IN_CITY_AMOUNT_SUM, new MyData(15, 13));
		locationMap.put(OTHER_EXPENSE_SUM, new MyData(15, 14));
		
		// 第六行
		locationMap.put(ITEMCODE, new MyData(18, 4));
		locationMap.put(REAL_AMOUNT, new MyData(18, 8));
		locationMap.put(UNIT, new MyData(18, 14));
		
		// 第七行
		locationMap.put(ITEMNAME, new MyData(19, 4));
		locationMap.put(AMOUNT_IN_LARGEWORDS, new MyData(19, 8));
		
		// 第八行
		locationMap.put(JOBNUM, new MyData(21, 4));
		
		// 第九行
		locationMap.put(CCB_ACCOUNT, new MyData(22, 8));
		locationMap.put(NAME_1, new MyData(22, 4));
				
		//第十行
		locationMap.put(TELE, new MyData(23, 4));
		locationMap.put(ICB_ACCOUNT, new MyData(23, 8));
		
		//起止时间等等
		for(int i=0;i<8;i++){
			locationMap.put(depar_month[i], new MyData(7+i, 0));
			locationMap.put(depar_date[i], new MyData(7+i, 1));
			locationMap.put(arri_month[i], new MyData(7+i, 2));
			locationMap.put(arri_date[i], new MyData(7+i, 3));	
			locationMap.put(depar_desti[i], new MyData(7+i, 4));	
			locationMap.put(travelling_expense[i], new MyData(7+i, 5));
			locationMap.put(car_fare[i], new MyData(7+i, 6));
			locationMap.put(accomodate_number[i], new MyData(7+i, 7));
			locationMap.put(accomodate_standard[i], new MyData(7+i, 8));
			locationMap.put(accomodate_amount[i], new MyData(7+i, 9));
			locationMap.put(meal_number[i], new MyData(7+i, 10));
			locationMap.put(meal_amount[i], new MyData(7+i, 11));
			locationMap.put(travel_in_city_number[i], new MyData(7+i, 12));
			locationMap.put(travel_in_city_amount[i], new MyData(7+i, 13));
			locationMap.put(other_expense[i], new MyData(7+i, 14));
		}
	}
}
