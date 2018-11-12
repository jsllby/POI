import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 测试类
 * @author fangzhiyang
 *
 */

public class Test {
    
	public static void main(String[] args) throws Exception {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
		//sheet页码
		int i=5;   
		Map<String, String> textMap = new HashMap<>();
		String fileName = "";
		// 工号
		String jobnum = "84122070";
		if(i==0){
			String realAmount = "1245";    //数字金额
			String itemcode = "87654321";    //项目代码
			String college = "1";    //单位
			String name = "1";    //姓名
			String CCB_account = "1";    //建行
			String ICB_account = "1";    //工行
			String reason = "1";    //事由
			String tele = "1";    //电话
			String itemname = "1";    //项目名称
			String payee = "1";    //收款单位
			String bank = "1";    //银行
			String account = "1";    //账号
			String agent = "1";    //经办人
			String receipt_num = "1";    //发票张数
			

			// 第一行
			textMap.put(LocationMap0.COLLEGE, college);
			textMap.put(LocationMap0.NAME, name);
			textMap.put(LocationMap0.CCB_ACCOUNT, CCB_account);
			textMap.put(LocationMap0.ICB_ACCOUNT, ICB_account);
			textMap.put(LocationMap0.REASON, reason);
			// 第二行
			textMap.put(LocationMap0.JOBNUM, jobnum);
			textMap.put(LocationMap0.TELE, tele);
			textMap.put(LocationMap0.ITEMCODE, itemcode);
			textMap.put(LocationMap0.ITEMNAME, itemname);
			// 第三行
			
			textMap.put(LocationMap0.REAL_AMOUNT, realAmount);
			textMap.put(LocationMap0.AMOUNT, "¥ " + ExcelUtil.digitUppercase(realAmount));

		    String digit = ExcelUtil.fillTextMapBySevenDigit(textMap.get(LocationMap0.REAL_AMOUNT), textMap);
            char[] charArray = digit.toCharArray();
	        textMap.put(LocationMap0.WAN, String.valueOf(charArray[0]));
		    textMap.put(LocationMap0.QIAN, String.valueOf(charArray[1]));
		    textMap.put(LocationMap0.BAI, String.valueOf(charArray[2]));
		    textMap.put(LocationMap0.SHI, String.valueOf(charArray[3]));
		    textMap.put(LocationMap0.YUAN, String.valueOf(charArray[4]));
		    textMap.put(LocationMap0.JIAO, String.valueOf(charArray[5]));
		    textMap.put(LocationMap0.FEN, String.valueOf(charArray[6]));
		    textMap.put(LocationMap0.DATE, sdf.format(new Date()));
		    // 第四行
    		textMap.put(LocationMap0.PAYEE, payee);
    		textMap.put(LocationMap0.BANK, bank);
  	     	textMap.put(LocationMap0.ACCOUNT, account);
  		    textMap.put(LocationMap0.AGENT, agent+"   "+tele);
  		    textMap.put(LocationMap0.RECEIPT_NUM, receipt_num);
		}
		else if(i==1){
			String realAmount = "1245";    //数字金额
			String itemcode = "87654321";    //项目代码
			String college = "1";    //单位
			String name = "1";    //姓名
			String CCB_account = "1";    //建行
			String ICB_account = "1";    //工行
			String reason = "1";    //事由
			String tele = "1";    //电话
			String itemname = "1";    //项目名称
			String payee = "1";    //收款单位
			String bank = "1";    //银行
			String account = "1";    //账号
			String agent = "1";    //经办人
			String receipt_num = "1";    //发票张数

//			// 第一行
			textMap.put(LocationMap1.COLLEGE, college);
			textMap.put(LocationMap1.NAME, name);
			textMap.put(LocationMap1.CCB_ACCOUNT, CCB_account);
			textMap.put(LocationMap1.ICB_ACCOUNT, ICB_account);
			textMap.put(LocationMap1.REASON, reason);
			// 第二行
			textMap.put(LocationMap1.JOBNUM, jobnum);
			textMap.put(LocationMap1.TELE, tele);
			textMap.put(LocationMap1.ITEMCODE, itemcode);
			textMap.put(LocationMap1.ITEMNAME, itemname);
			// 第三行
			
			textMap.put(LocationMap1.REAL_AMOUNT, realAmount);
			textMap.put(LocationMap1.AMOUNT, "¥ " + ExcelUtil.digitUppercase(realAmount));

		    String digit = ExcelUtil.fillTextMapBySevenDigit(textMap.get(LocationMap0.REAL_AMOUNT), textMap);
            char[] charArray = digit.toCharArray();
	        textMap.put(LocationMap1.WAN, String.valueOf(charArray[0]));
		    textMap.put(LocationMap1.QIAN, String.valueOf(charArray[1]));
		    textMap.put(LocationMap1.BAI, String.valueOf(charArray[2]));
		    textMap.put(LocationMap1.SHI, String.valueOf(charArray[3]));
		    textMap.put(LocationMap1.YUAN, String.valueOf(charArray[4]));
		    textMap.put(LocationMap1.JIAO, String.valueOf(charArray[5]));
		    textMap.put(LocationMap1.FEN, String.valueOf(charArray[6]));
		    textMap.put(LocationMap1.DATE, sdf.format(new Date()));
		    // 第四行
    		textMap.put(LocationMap1.PAYEE, payee);
    		textMap.put(LocationMap1.BANK, bank);
  	     	textMap.put(LocationMap1.ACCOUNT, account);
  		    textMap.put(LocationMap1.AGENT, agent+"   "+tele);
  		    textMap.put(LocationMap1.RECEIPT_NUM, receipt_num);
	

	        //冲账
			// 第一行
  		    textMap.put(LocationMap1.COLLEGE_2, college);
			textMap.put(LocationMap1.NAME_2, name);
			textMap.put(LocationMap1.CCB_ACCOUNT_2, CCB_account);
			textMap.put(LocationMap1.ICB_ACCOUNT_2, ICB_account);
			textMap.put(LocationMap1.REASON_2, reason);
			// 第二行
			textMap.put(LocationMap1.JOBNUM_2, jobnum);
			textMap.put(LocationMap1.TELE_2, tele);
			textMap.put(LocationMap1.ITEMCODE_2, itemcode);
			textMap.put(LocationMap1.ITEMNAME_2, itemname);
			// 第三行		
			textMap.put(LocationMap1.AMOUNT_2, "¥ " + ExcelUtil.digitUppercase(realAmount));
			textMap.put(LocationMap1.WAN_2, String.valueOf(charArray[0]));
			textMap.put(LocationMap1.QIAN_2, String.valueOf(charArray[1]));
			textMap.put(LocationMap1.BAI_2, String.valueOf(charArray[2]));
			textMap.put(LocationMap1.SHI_2, String.valueOf(charArray[3]));
			textMap.put(LocationMap1.YUAN_2, String.valueOf(charArray[4]));
			textMap.put(LocationMap1.JIAO_2, String.valueOf(charArray[5]));
			textMap.put(LocationMap1.FEN_2, String.valueOf(charArray[6]));
    		// 第四行
			textMap.put(LocationMap1.AGENT_2, agent+"   "+tele);
			textMap.put(LocationMap1.RECEIPT_NUM_2, receipt_num);
			textMap.put(LocationMap1.DATE_2, sdf.format(new Date()));		
	}
		else if(i==2){
			String realAmount = "123467.5";	  //数字金额
			String itemcode = "87654321";	  //ITEMCODE
			String college = "2";	  //单位
			String CCB_amount = "2";	  //建行
			String ICB_amount = "2";	  //工行
			String reason = "2";	  //事由
			String name = "2";	  //姓名
			String tele = "2";	  //电话
			String itemname = "2";	  //项目名称
			String payee = "2";	  //收款单位
			String bank = "2";	  //银行
			String account = "2";	  //账号
			String receipt_num = "2";	  //发票张数
			String agent = "2";	  //经办人

			// 第一行
			textMap.put(LocationMap2.COLLEGE, college);
	 		textMap.put(LocationMap2.NAME, name);
			textMap.put(LocationMap2.CCB_ACCOUNT, CCB_amount);
			textMap.put(LocationMap2.ICB_ACCOUNT, ICB_amount);
			textMap.put(LocationMap2.REASON, reason);
			// 第二行
			textMap.put(LocationMap2.JOBNUM, jobnum);
			textMap.put(LocationMap2.TELE, tele);
			textMap.put(LocationMap2.ITEMCODE, itemcode);
			textMap.put(LocationMap2.ITEMNAME, itemname);
			// 第三行
			
			textMap.put(LocationMap2.REAL_AMOUNT, realAmount);
			textMap.put(LocationMap2.AMOUNT, "¥ " + ExcelUtil.digitUppercase(realAmount));
			String digit = ExcelUtil.fillTextMapByElevenDigit(textMap.get(LocationMap2.REAL_AMOUNT), textMap);
			char[] charArray = digit.toCharArray();
			textMap.put(LocationMap2.YI, String.valueOf(charArray[0]));
			textMap.put(LocationMap2.QIANWAN, String.valueOf(charArray[1]));
			textMap.put(LocationMap2.BAIWAN, String.valueOf(charArray[2]));
			textMap.put(LocationMap2.SHIWAN, String.valueOf(charArray[3]));
			textMap.put(LocationMap2.WAN, String.valueOf(charArray[4]));
			textMap.put(LocationMap2.QIAN, String.valueOf(charArray[5]));
			textMap.put(LocationMap2.BAI, String.valueOf(charArray[6]));
			textMap.put(LocationMap2.SHI, String.valueOf(charArray[7]));
			textMap.put(LocationMap2.YUAN, String.valueOf(charArray[8]));
			textMap.put(LocationMap2.JIAO, String.valueOf(charArray[9]));
			textMap.put(LocationMap2.FEN, String.valueOf(charArray[10]));

			// 第四行
			textMap.put(LocationMap2.PAYEE, payee);
			textMap.put(LocationMap2.BANK, bank);
			textMap.put(LocationMap2.ACCOUNT, account);
			textMap.put(LocationMap2.AGENT, agent+"  "+tele);
			textMap.put(LocationMap2.RECEIPT_NUM , receipt_num);
			textMap.put(LocationMap2.DATE, sdf.format(new Date()));

			//冲账
			// 第一行
			textMap.put(LocationMap2.COLLEGE_1, college);
			textMap.put(LocationMap2.NAME_1, name);
			textMap.put(LocationMap2.CCB_ACCOUNT_1, CCB_amount);
			textMap.put(LocationMap2.ICB_ACCOUNT_1, ICB_amount);
			textMap.put(LocationMap2.REASON_1, reason);
			// 第二行
			textMap.put(LocationMap2.JOBNUM_1, jobnum);
			textMap.put(LocationMap2.TELE_1, tele);
			textMap.put(LocationMap2.ITEMCODE_1, itemcode);
			textMap.put(LocationMap2.ITEMNAME_1, itemname);
			// 第三行
			
			textMap.put(LocationMap2.AMOUNT_1, "¥ " + ExcelUtil.digitUppercase(realAmount));
			
			textMap.put(LocationMap2.YI_1, String.valueOf(charArray[0]));
			textMap.put(LocationMap2.QIANWAN_1, String.valueOf(charArray[1]));
			textMap.put(LocationMap2.BAIWAN_1, String.valueOf(charArray[2]));
			textMap.put(LocationMap2.SHIWAN_1, String.valueOf(charArray[3]));
			textMap.put(LocationMap2.WAN_1, String.valueOf(charArray[4]));
			textMap.put(LocationMap2.QIAN_1, String.valueOf(charArray[5]));
			textMap.put(LocationMap2.BAI_1, String.valueOf(charArray[6]));
			textMap.put(LocationMap2.SHI_1, String.valueOf(charArray[7]));
			textMap.put(LocationMap2.YUAN_1, String.valueOf(charArray[8]));
			textMap.put(LocationMap2.JIAO_1, String.valueOf(charArray[9]));
			textMap.put(LocationMap2.FEN_1, String.valueOf(charArray[10]));

			// 第四行
			textMap.put(LocationMap2.AGENT_1, agent+"   "+tele);
			textMap.put(LocationMap2.RECEIPT_NUM_1, receipt_num);
			textMap.put(LocationMap2.DATE_1, sdf.format(new Date()));		
		}
		else if(i==3){
			String realAmount = "123467.5";	  //数字金额
			String itemcode = "87654321";	  //ITEMCODE
			String college = "2";	  //单位
			String CCB_amount = "2";	  //建行
			String ICB_amount = "2";	  //工行
			String reason = "2";	  //事由
			String name = "2";	  //姓名
			String tele = "2";	  //电话
			String itemname = "2";	  //项目名称
			String payee = "2";	  //收款单位
			String bank = "2";	  //银行
			String account = "2";	  //账号
			String receipt_num = "2";	  //发票张数
			String agent = "2";	  //经办人

			// 第一行
			textMap.put(LocationMap2.COLLEGE, college);
	 		textMap.put(LocationMap2.NAME, name);
			textMap.put(LocationMap2.CCB_ACCOUNT, CCB_amount);
			textMap.put(LocationMap2.ICB_ACCOUNT, ICB_amount);
			textMap.put(LocationMap2.REASON, reason);
			// 第二行
			textMap.put(LocationMap2.JOBNUM, jobnum);
			textMap.put(LocationMap2.TELE, tele);
			textMap.put(LocationMap2.ITEMCODE, itemcode);
			textMap.put(LocationMap2.ITEMNAME, itemname);
			// 第三行
			
			textMap.put(LocationMap2.REAL_AMOUNT, realAmount);
			textMap.put(LocationMap2.AMOUNT, "¥ " + ExcelUtil.digitUppercase(realAmount));
			String digit = ExcelUtil.fillTextMapByElevenDigit(textMap.get(LocationMap2.REAL_AMOUNT), textMap);
			char[] charArray = digit.toCharArray();
			textMap.put(LocationMap2.YI, String.valueOf(charArray[0]));
			textMap.put(LocationMap2.QIANWAN, String.valueOf(charArray[1]));
			textMap.put(LocationMap2.BAIWAN, String.valueOf(charArray[2]));
			textMap.put(LocationMap2.SHIWAN, String.valueOf(charArray[3]));
			textMap.put(LocationMap2.WAN, String.valueOf(charArray[4]));
			textMap.put(LocationMap2.QIAN, String.valueOf(charArray[5]));
			textMap.put(LocationMap2.BAI, String.valueOf(charArray[6]));
			textMap.put(LocationMap2.SHI, String.valueOf(charArray[7]));
			textMap.put(LocationMap2.YUAN, String.valueOf(charArray[8]));
			textMap.put(LocationMap2.JIAO, String.valueOf(charArray[9]));
			textMap.put(LocationMap2.FEN, String.valueOf(charArray[10]));

			// 第四行
			textMap.put(LocationMap2.PAYEE, payee);
			textMap.put(LocationMap2.BANK, bank);
			textMap.put(LocationMap2.ACCOUNT, account);
			textMap.put(LocationMap2.AGENT, agent+"  "+tele);
			textMap.put(LocationMap2.RECEIPT_NUM , receipt_num);
			textMap.put(LocationMap2.DATE, sdf.format(new Date()));
		}
		else if(i==4){
			String Approver = "四川大学";			//批准单位
			String Reason = "差旅";			//事由
			String Name = "张三";			//姓名
			String Title = "讲师";		//职务职称
			String itemcode = "12345";		//项目编号
			String Assistance = "是";		//是否报销补助
			String BillAmount = "3";		//单据张数
			String Unit = "四川大学";		//派出单位
			String itemname = "叫什么呢";	//项目名称
			String CCB_amount = "8888";   //建行卡号
			String ICB_amount = "8888";		//工行卡号
			String tele = "12580";		//电话
			
			int Tirpnum = 5;   //  <=8    
			
			String[] Depar_month = new String[] {"1","2","3","4","5"};    //出发月
			String[] Depar_date = new String[] {"11","12","13","14","15"};		//出发日
			String[] Arri_month = new String[] {"1","2","3","4","5"};		//到达月
			String[] Arri_date = new String[] {"12","13","14","15","16"};		//到达日
			String[] Departure = new String[] {"A","B","C","D","E"};		//出发地
			String[] Destination = new String[] {"AA","BB","CC","DD","EE"};		//目的地
			int[] Travelling_expense = new int[] {1,2,3,4,5};		//车船费飞机票
			int[] Car_fare = new int[] {11,22,33,44,55};		//车费
			String[] Accomodate_number = new String[] {"1","1","1","1","1"};		//住宿人数
			String[] Accomodate_standard = new String[] {"100","100","100","100","100"};		//住宿标准
			int[] Accomodate_amount = new int[] {200,200,200,200,200};		//住宿金额
			int[] Other_expense = new int[] {100,100,100,100,100};
			
			int Travelling_expense_sum = 0;		//船费飞机票合计
			int Car_fare_sum = 0;		//车费合计
			int Accomodate_amount_sum = 0;		//住宿金额合计
			int Other_expense_sum = 0;		//其他费用合计
			int Amount = 0;		//合计
	
			// 第一行
			textMap.put(LocationMap4.APPROVER, Approver);
			textMap.put(LocationMap4.REASON, Reason);
			
			// 第二行
			textMap.put(LocationMap4.NAME, Name);
			
			// 第三行		
			textMap.put(LocationMap4.TITLE, Title);
			
			//第四行
			textMap.put(LocationMap4.ASSISTANCE, Assistance);
			
			//第五行
			textMap.put(LocationMap4.BILL_AMOUNT, BillAmount);
			
			//第六行		
			textMap.put(LocationMap4.ITEMCODE, itemcode);
			textMap.put(LocationMap4.UNIT, Unit);
			
			//第七行
			textMap.put(LocationMap4.ITEMNAME, itemname);
			
			//第八行
			textMap.put(LocationMap4.JOBNUM, jobnum);
			
			//第九行
			textMap.put(LocationMap4.NAME_1, Name);
			textMap.put(LocationMap4.CCB_ACCOUNT, CCB_amount);
			
			//第十行
			textMap.put(LocationMap4.TELE, tele);
			textMap.put(LocationMap4.ICB_ACCOUNT, ICB_amount);
			
			//行程
			for(int j=0;j<Tirpnum;j++){
				textMap.put(LocationMap4.depar_month[j], Depar_month[j]);
				textMap.put(LocationMap4.depar_date[j], Depar_date[j]);
				textMap.put(LocationMap4.arri_month[j], Arri_month[j]);
				textMap.put(LocationMap4.arri_date[j], Arri_date[j]);
				textMap.put(LocationMap4.depar_desti[j], Departure[j]+"——"+Destination[j]);
				textMap.put(LocationMap4.travelling_expense[j], "" + Travelling_expense[j]);
				textMap.put(LocationMap4.car_fare[j], "" + Car_fare[j]);
				textMap.put(LocationMap4.accomodate_number[j],Accomodate_number[j]);
				textMap.put(LocationMap4.accomodate_standard[j],Accomodate_standard[j]);
				textMap.put(LocationMap4.accomodate_amount[j], "" + Accomodate_amount[j]);
				textMap.put(LocationMap4.other_expense[j], "" + Other_expense[j]);
				Travelling_expense_sum += Travelling_expense[j];
				Car_fare_sum += Car_fare[j];
				Accomodate_amount_sum += Accomodate_amount[j];
				Other_expense_sum += Other_expense[j];
			}
			
			textMap.put(LocationMap4.TRAVELLING_EXPENSE_SUM, "" + Travelling_expense_sum);
			textMap.put(LocationMap4.CAR_FARE_SUM, "" + Car_fare_sum);
			textMap.put(LocationMap4.ACCOMODATE_AMOUNT_SUM, "" + Accomodate_amount_sum);
			
			Amount = Travelling_expense_sum + Car_fare_sum + Accomodate_amount_sum + Other_expense_sum;
			
			textMap.put(LocationMap4.AMOUNT, "" + Amount);	
			textMap.put(LocationMap4.AMOUNT_IN_LARGEWORDS, "¥ " + ExcelUtil.digitUppercase("" + Amount));
		}
		else if(i==5){
			String payer = "计算机学院";   //付款单位
			String payee = "文新学院";   //收款单位
			String topic = "大创项目";   //课题名称
			String department = "123";	//部门代码
			String itemcode = "0055";	//项目代码
			String reason = "给我钱";		//事由
			int amount = 10;	//数量
			int price = 100;	//单价
			String realAmount = "" + amount*price;   //金额
			String payer_transactor = "小雪花";   //付款单位经办人
			String payee_transactor = "小雪fa";		//收款单位经办人
			String transactor_jobnum_1 = "8732845";	//付款单位经办人工作证号
			String transactor_jobnum_2 = "8523227";	//收款单位经办人工作证号
			String tele_1 = "10086";	//付款单位经办人电话
			String tele_2 = "100861";   //收款单位经办人电话

//			// 第一行
		    textMap.put(LocationMap5.PAYER, payer);
			textMap.put(LocationMap5.PAYEE, payee);
			//第二行
			textMap.put(LocationMap5.TOPIC_1, topic);
			textMap.put(LocationMap5.TOPIC_2, topic);
			textMap.put(LocationMap5.DEPARTMENT_1,department);
			textMap.put(LocationMap5.DEPARTMENT_2,department);
			textMap.put(LocationMap5.ITEMCODE_1,itemcode);
			textMap.put(LocationMap5.ITEMCODE_2,itemcode);

			// 第三行
			textMap.put(LocationMap5.REASON, reason);
			//第四行
			textMap.put(LocationMap5.SIZE, "数量  "+ amount+"  单价    "+ price);
			//第五行
			textMap.put(LocationMap5.AMOUNT_IN_SMALLWORDS,realAmount);
			textMap.put(LocationMap5.AMOUNT_IN_LARGEWORDS,ExcelUtil.digitUppercase(realAmount));
			//第六行
			textMap.put(LocationMap5.PAYER_TRANSACTOR,payer_transactor);
			textMap.put(LocationMap5.PAYEE_TRANSACTOR,payee_transactor);
			
			textMap.put(LocationMap5.TRANSACTOR_JOBNUM_1,transactor_jobnum_1);
			textMap.put(LocationMap5.TRANSACTOR_JOBNUM_2,transactor_jobnum_2);
			textMap.put(LocationMap5.TELE_1,tele_1);
			textMap.put(LocationMap5.TELE_2,tele_2);
			textMap.put(LocationMap5.REAL_AMOUNT,realAmount);

			textMap.put(LocationMap5.DATE, sdf.format(new Date()));
			
			if (itemcode != null && !itemcode.equals(""))
				fileName = itemcode + "报账单.xls";
			else
				fileName = "报账单.xls";	
		}

		if(i<5){
			if (jobnum != null && !jobnum.equals(""))
				fileName = jobnum + "报账单.xls";
			else
				fileName = "报账单.xls";
		}
		ExcelUtil toScufdExcel = new ExcelUtil(textMap,jobnum,i);
		toScufdExcel.writeExcel(fileName);
	}
}
