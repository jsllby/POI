package com.ruite.util.poi;

import java.util.Collection;
import java.util.HashMap;

public class LocationMap {
	public static final String DATE = "日期";
	public static final String REAL_AMOUNT = "数字金额";
	public static final String ITEMCODE = "项目代码";
	
	protected static HashMap<String, MyData> locationMap = new HashMap<>();

	public MyData getData(String name) {
		return locationMap.get(name);
	}

	public Collection<MyData> getValues() {
		return locationMap.values();
	}

}
