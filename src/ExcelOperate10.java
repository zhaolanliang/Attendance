
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelOperate10 {
	public static void main(String[] args) throws ParseException, IOException{
		
//创建新工作表 
		   HSSFWorkbook wb = new HSSFWorkbook();
		   HSSFCellStyle style = wb.createCellStyle();
		   style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	        HSSFSheet table = wb.createSheet("统计表");	
			   HSSFRow Crow = table.createRow((int)0);	 
			   HSSFCell s_cell = Crow.createCell((short)0);
			   s_cell.setCellValue("姓名");
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 1);  
			   s_cell.setCellValue("平均上班时间");  
			   s_cell.setCellStyle(style);  
			   s_cell = Crow.createCell((short) 2);  
			   s_cell.setCellValue("平均下班时间");  
			   s_cell.setCellStyle(style);  
			   s_cell = Crow.createCell((short) 3);  
			   s_cell.setCellValue("月出勤天数");  
			   s_cell.setCellStyle(style); 
			   s_cell = Crow.createCell((short) 4);  
			   s_cell.setCellValue("负激励金总额");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 5);  
			   s_cell.setCellValue("日均出勤时长");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 6);  
			   s_cell.setCellValue("迟到0-15分钟");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 7);  
			   s_cell.setCellValue("迟到16-30分钟");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 8);  
			   s_cell.setCellValue("迟到31-60分钟");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 9);  
			   s_cell.setCellValue("早退0-15分钟");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 10);  
			   s_cell.setCellValue("早退16-30分钟");  
			   s_cell.setCellStyle(style);
			   s_cell = Crow.createCell((short) 11);  
			   s_cell.setCellValue("早退31-60分钟");  
			   s_cell.setCellStyle(style);
			   
			   
		   
//导入数据
		FileInputStream finput = new FileInputStream("D://Attendance//原始打卡记录表.xls" );
		POIFSFileSystem fs = new POIFSFileSystem( finput );
		HSSFWorkbook hs = new HSSFWorkbook(fs);
		HSSFSheet readSheet =null;
//处理数据
		int day = 0;
		readSheet = hs.getSheetAt(0);
		String[][] result = null;
		try {
			result = getData(readSheet,1);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		 
		SimpleDateFormat time = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date date = null;	
		
		Map<String,List> map = new HashMap<String,List>();
		
		for(int rowname=0; rowname<result.length;rowname++){
			
			if(map.containsKey(result[rowname][3])){
				map.get(result[rowname][3]).add(result[rowname][4]);
			}else{
				List<Object> L = new ArrayList<Object>();
				map.put(result[rowname][3], L);
				map.get(result[rowname][3]).add(result[rowname][4]);
			}
		}
		
		
//逻辑处理
		Iterator<?> maps = map.entrySet().iterator(); 
		int i=0;
		while(maps.hasNext()){
			
			Map.Entry<String, List> LIST=(Map.Entry<String, List>)maps.next(); 
			List values = LIST.getValue();
			
			HSSFSheet sheet = wb.createSheet(LIST.getKey());
			HSSFRow row = sheet.createRow((int)0);
			HSSFCell cell = row.createCell((short)0);
			cell.setCellValue("姓名");
			cell.setCellStyle(style);
			cell = row.createCell((short) 1);  
	        cell.setCellValue("上班时间");  
	        cell.setCellStyle(style);  
	        cell = row.createCell((short) 2);  
	        cell.setCellValue("下班时间");  
	        cell.setCellStyle(style);  
	        cell = row.createCell((short) 3);  
	        cell.setCellValue("上下班状态");  
	        cell.setCellStyle(style); 
	        cell = row.createCell((short) 4);  
	        cell.setCellValue("负激励金");  
	        cell.setCellStyle(style);
	
			double workTime = 0;
			int daycount = 0;//月上班天数
			double sumdaytime = 0;//月上班时间
			double onworkTime = 0;//上班时间
			double outworkTime = 0;//下班时间
			double breaklong = 0;//迟到及早退时间
			int breaklong1 =0;
			int breaklong2 = 0;
			int money = 0;//负激励金
			int late_15 = 0;
			int late_30 = 0;
			int late_60 = 0;
			int leave_15 = 0;
			int leave_30 = 0;
			int leave_60 = 0;
	//区分打卡时间是上班还是下班
			Map<Integer,Date> map1 = new HashMap<Integer,Date>();
			Map<Integer,Date> map2 = new HashMap<Integer,Date>();
			
			for(int p=0; p<values.size();p++){
				date = time.parse((String) values.get(p));
				day = date.getDate();
				if(date.getHours() < 12){
					map1.put(day, date);
				}else{
					map2.put(day, date);
				}
			}
			Iterator<?> onwork = map1.entrySet().iterator(); 
			Iterator<?> breakwork = map2.entrySet().iterator();
			System.out.println("姓名"+"\t"+"上班时间"+"\t\t\t"+"下班时间"+"\t\t\t"+"上下班状态"+"\t\t"+"备注");
		  	System.out.println("***********************************************************************************************************");
		  	int q = -1;
		  	while(onwork.hasNext()){ 
				   q++;
				   Map.Entry<Integer, Date> m=(Map.Entry<Integer, Date>)onwork.next(); 
				   Object key1 = m.getKey();
				   Date value1 = m.getValue();
				   SimpleDateFormat time1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				   String datetime1  = time1.format(value1);
				   int H1 = value1.getHours();
				   int M1 = value1.getMinutes();
				   String Hour1 = null;
				   String MM1 = null;
				   String Hour2 = null;
				   String MM2 = null;
				   row = sheet.createRow((int)q);
				   if(H1<10){
					   Hour1 = '0'+String.valueOf(H1);
				   }else{
					   Hour1 = String.valueOf(H1);
				   }
				   if(M1<10){
					   MM1 = '0'+String.valueOf(M1);
				   }else{
					   MM1 = String.valueOf(M1);
				   }
				   row.createCell((short)0).setCellValue(LIST.getKey());
				   row.createCell((short)1).setCellValue(datetime1);
				   System.out.print(LIST.getKey()+"\t");
				   System.out.print( m.getKey() +"号上班时间："+ Hour1 + ":" + MM1+"\t\t");
				   
				for(Entry<Integer, Date> entry : map2.entrySet()){
					Object key2 = entry.getKey();
					Date value2 = (Date)entry.getValue();
					SimpleDateFormat time2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
					String datetime2  = time2.format(value2);
					int H2 = value2.getHours();
					int M2 = value2.getMinutes();
					if(H2<10){
						   Hour2 = '0'+String.valueOf(H2);
					   }else{
						   Hour2 = String.valueOf(H2);
					   }
					   if(M2<10){
						   MM2 = '0'+String.valueOf(M2);
					   }else{
						   MM2 = String.valueOf(M2);
					   }
				   if(map2.containsKey(key1)){
					   if(key1.equals(key2)){
						  row.createCell((short)2).setCellValue(datetime2);
						   System.out.print(entry.getKey()+"号下班时间：" + Hour2+":" + MM2 +"\t\t");
						   //9点之前到，18点之前后班，正常上下班
						   if(((H1<9)||(H1==9 && M1 == 0))&& (H2>=18)){
					    		workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
					    		sumdaytime = sumdaytime+workTime;
					    		onworkTime = onworkTime+H1*60+M1;
					    		outworkTime = outworkTime+H2*60+M2;
					    		row.createCell((short)3).setCellValue((String)"正常上下班");
					    		row.createCell((short)4).setCellValue(0);
					    		System.out.println("正常上下班");
					    		daycount++;
					    		//9点之前到，18点之前下班，早退
					    	}else if(((H1<9)||(H1==9 && M1 == 0))&& ((H2<18))){
					    		if(H2<17){
					    			row.createCell((short)3).setCellValue((String)"事假");
					    			row.createCell((short)4).setCellValue(0);
					    			System.out.println("事假");
					    		}else{
						    		workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
						    		sumdaytime = sumdaytime+workTime;
						    		onworkTime = onworkTime+H1*60+M1;
						    		outworkTime = outworkTime+H2*60+M2;
						    		breaklong = 60-M2;
						    		row.createCell((short)3).setCellValue((String)("早退"+breaklong+"分钟"));
						    		System.out.print("早退"+breaklong+"分钟"+"\t\t");
						    		if(breaklong<=15){
						    		    leave_15++;
						    			money = money-20;
						    			row.createCell((short)4).setCellValue(-20);
						    			System.out.println("-20");
						    		}else if(breaklong<=30){
						    			leave_30++;
						    			money=money-40;
						    			row.createCell((short)4).setCellValue(-40);
						    			System.out.println("-40");
						    		}else{
						    			leave_60++;
						    			money = money-80;
						    			row.createCell((short)4).setCellValue(-80);
						    			System.out.println("-80");
						    		}
						    		daycount++;
					    		}
					    		//9:01-9:59上班，18点之前下班,早退
					    	}else if((H1==9)&&(M1>0)&&(H2<18)){
					    		workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
					    		if(H2 <17){
					    			row.createCell((short)3).setCellValue((String)"事假");
					    			row.createCell((short)4).setCellValue(0);
					    			System.out.println("事假");
					    		}else{
					    			sumdaytime = sumdaytime+workTime;
					    			onworkTime = onworkTime+H1*60+M1;
					    			outworkTime = outworkTime+H2*60+M2;
					    			breaklong = 60-M2;
					    			row.createCell((short)3).setCellValue((String)("早退"+breaklong+"分钟"));
					    			System.out.print("早退"+breaklong+"分钟"+"\t\t");
					    			if(breaklong<=15){
						    			leave_15++;
					    				money = money-20;
						    			row.createCell((short)4).setCellValue(-20);
						    			System.out.println("-20");
						    		}else if(breaklong<=30){
						    			leave_30++;
						    			money=money-40;
						    			row.createCell((short)4).setCellValue(-40);
						    			System.out.println("-40");
						    		}else{
						    			leave_60++;
						    			money = money-80;
						    			row.createCell((short)4).setCellValue(-80);
						    			System.out.println("-80");
						    		}
					    			daycount++;
					    		}
					    		//9:01-9:59上班，18：00-18：59下班,迟到
					    	}else if((H1==9&&M1>0)&&(H2==18)){
					    		workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
					    			sumdaytime = sumdaytime+workTime;
					    			onworkTime = onworkTime+H1*60+M1;
					    			outworkTime = outworkTime+H2*60+M2;
					    			breaklong = M1;
					    			row.createCell((short)3).setCellValue((String)("迟到"+breaklong+"分钟"));
					    			System.out.print("迟到"+breaklong+"分钟"+"\t\t");
					    			if(breaklong<=15){
					    				late_15++;
						    			money = money-20;
						    			row.createCell((short)4).setCellValue(-20);
						    			System.out.println("-20");
						    		}else if(breaklong<=30){
						    			late_30++;
						    			money=money-40;
						    			row.createCell((short)4).setCellValue(-40);
						    			System.out.println("-40");
						    		}else{
						    			late_60++;
						    			money = money-80;
						    			row.createCell((short)4).setCellValue(-80);
						    			System.out.println("-80");
						    		}
					    			daycount++;

						    		//9:01-9:59上班，19：00之后下班,正常上下班
						    	}else if((H1==9&&M1>0)&&(H2>=19)){
						    		workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
						    			sumdaytime = sumdaytime+workTime;
						    			onworkTime = onworkTime+H1*60+M1;
						    			outworkTime = outworkTime+H2*60+M2;
						    			row.createCell((short)3).setCellValue((String)("正常上下班"));
						    			row.createCell((short)4).setCellValue(0);
						    			System.out.println("正常上下班");
						    			daycount++;
						    		
						    		//10点之后到，19点之前下班，迟到+早退
						    	}else if((H1>=10)&&H2<19){
						    		    breaklong1 = M1;
						    		    breaklong2 = 60-M2;
						    			breaklong = M1+60-M2;
						    			if(breaklong>=60){
						    				row.createCell((short)3).setCellValue((String)("事假"));
						    				row.createCell((short)4).setCellValue(0);
						    				System.out.println("事假");
							    		}else if(H1==10&&M1==0){
							    			workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
								    		sumdaytime = sumdaytime+workTime;
								    		onworkTime = onworkTime+H1*60+M1;
								    		outworkTime = outworkTime+H2*60+M2;
								    		breaklong = 60-M2;
								    		row.createCell((short)3).setCellValue((String)"早退"+breaklong+"分钟");
								    		System.out.print("早退"+breaklong+"分钟"+"\t\t");
								    		if(breaklong<=15){
								    			leave_15++;
								    			money = money-20;
								    			row.createCell((short)4).setCellValue(-20);
								    			System.out.println("-20");
								    		}else if(breaklong<=30){
								    			leave_30++;
								    			money=money-40;
								    			row.createCell((short)4).setCellValue(-40);
								    			System.out.println("-40");
								    		}else{
								    			leave_60++;
								    			money = money-80;
								    			row.createCell((short)4).setCellValue(-80);
								    			System.out.println("-80");
								    		}
								    		daycount++;
							    		}else{
							    			sumdaytime = sumdaytime+workTime;
							    			onworkTime = onworkTime+H1*60+M1;
							    			outworkTime = outworkTime+H2*60+M2;
							    			int moneytemporary = 0;
							    			row.createCell((short)3).setCellValue((String)("迟到"+breaklong1+"分钟，"+"早退"+breaklong2+"分钟"));
							    			if(breaklong1<=15){
							    				moneytemporary = moneytemporary-20;
								    			late_15++;
								    			System.out.print("迟到"+breaklong1+"分钟");
								    		}else if(breaklong1<=30){
								    			moneytemporary = moneytemporary-40;
								    			late_30++;
								    			System.out.print("迟到"+breaklong1+"分钟");
								    		}else{
								    			moneytemporary = moneytemporary-80;
								    			late_60++;
								    			System.out.print("迟到"+breaklong1+"分钟");
								    		}
							    			if(breaklong2<=15){
							    				moneytemporary = moneytemporary-20;
								    			leave_15++;
								    			System.out.print("早退"+breaklong2+"分钟");
								    		}else if(breaklong2<=30){
								    			moneytemporary = moneytemporary-40;
								    			leave_30++;
								    			System.out.print("早退"+breaklong2+"分钟");
								    		}else{
								    			moneytemporary = moneytemporary-80;
								    			leave_60++;
								    			System.out.print("早退"+breaklong2+"分钟");
								    		}
							    			money = money+moneytemporary;
							    			System.out.print("\t");
							    			if(breaklong<=15){
								    			row.createCell((short)4).setCellValue(moneytemporary);
								    			System.out.println(moneytemporary);
								    		}else if(breaklong<=30){
								    			row.createCell((short)4).setCellValue(moneytemporary);
								    			System.out.println(moneytemporary);
								    		}else{
								    			row.createCell((short)4).setCellValue(moneytemporary);
								    			System.out.println(moneytemporary);
								    		}
							    			daycount++;
							    		}
						    		//10点之后到，19点之后下班,上班迟到
						    	}else if(H1>=10&&H2>=19){
						    		if(H1>=11){
						    			row.createCell((short)3).setCellValue((String)("事假"));
						    			System.out.println("事假"+"\t\t");
						    		}else if((H1==10&&M1==0)&&(H2==19&&M2>=0)){
						    			workTime = (H2*60+M2)-(H1*60+M1)-1.5*60;
						    			sumdaytime = sumdaytime+workTime;
						    			onworkTime = onworkTime+H1*60+M1;
						    			outworkTime = outworkTime+H2*60+M2;
						    			row.createCell((short)3).setCellValue((String)("正常上下班"));
						    			row.createCell((short)4).setCellValue(0);
						    			System.out.println("正常上下班");
						    			daycount++;
						    		}else{
							    		sumdaytime = sumdaytime+workTime;
						    			onworkTime = onworkTime+H1*60+M1;
						    			outworkTime = outworkTime+H2*60+M2;
						    			breaklong = M1;
						    			row.createCell((short)3).setCellValue((String)("迟到"+breaklong+"分钟"));
						    			System.out.print("迟到"+breaklong+"分钟"+"\t\t");
						    			if(breaklong<=15){
						    				late_15++;
							    			money = money-20;
							    			row.createCell((short)4).setCellValue(-20);
							    			System.out.println("-20");
							    		}else if(breaklong<=30){
							    			late_30++;
							    			money=money-40;
							    			row.createCell((short)4).setCellValue(-40);
							    			System.out.println("-40");
							    		}else{
							    			late_60++;
							    			money = money-80;
							    			row.createCell((short)4).setCellValue(-80);
							    			System.out.println("-80");
							    		}
						    			daycount++;
					    			}
						    	}
						    	
						    	break;
						   }
					   }else{
						   row.createCell((short)2).setCellValue((String)"未打下班卡");
					    	System.out.print("未打下班卡"+"\t\t\t");
					    	row.createCell((short)3).setCellValue((String)"事假");
					    	row.createCell((short)4).setCellValue(0);
					    	System.out.println("事假");
					    	break;
					    }
					}
		  	}
				i++;//控制生成每个员工的上下班打卡信息表
				HSSFRow Crow1 = table.createRow((int)i);
				Crow1.createCell((short)0).setCellValue((String)LIST.getKey());
				DecimalFormat    df   = new DecimalFormat("######0.00");
				double dayavgtime = (sumdaytime/60)/daycount;
				String dt = df.format(dayavgtime);
		  	
				System.out.println("***********************************************************************************************************");
				   
				 int onwt =(int) ((onworkTime/60)/daycount);
				   int Om = (int)(((onworkTime/60)/daycount-onwt)*60);
				   String ONWT = null;
				   if(onwt<10){
					   ONWT = String.valueOf(onwt);
					   ONWT = "0"+ONWT;
				   }else{
					   ONWT = String.valueOf(onwt);
				   }
				   String OM = null;
				   if(Om<10){
					   OM = String.valueOf(Om);
					   OM = "0"+OM;
				   }else{
					   OM = String.valueOf(Om);
				   }
				   Crow1.createCell((short)1).setCellValue(ONWT+":"+OM);
				   System.out.print(LIST.getKey()+"\t");
				   System.out.print("平均上班时间："+ONWT+":"+OM+"\t\t");
				   int ouwt =(int) ((outworkTime/60)/daycount);
				   int Ou = (int)((((outworkTime/60)/daycount)-ouwt)*60);
				   String OUWT = null;
				   if(ouwt<10){
					   OUWT = String.valueOf(ouwt);
					   OUWT = "0"+OUWT;
				   }else{
					   OUWT = String.valueOf(ouwt);
				   }
				   String OU = null;
				   if(Ou<10){
					   OU = String.valueOf(Ou);
					   OU = "0"+OU;
				   }else{
					   OU = String.valueOf(Ou);
				   }
				   Crow1.createCell((short)2).setCellValue(OUWT+":"+OU);
				   System.out.print("平均下班时间："+OUWT+":"+OU+"\t\t");
				   Crow1.createCell((short)3).setCellValue(daycount);
				   System.out.print("月出勤天数："+daycount+"\t\t");
				   Crow1.createCell((short)4).setCellValue(money);
				   System.out.println("总计："+money);
				   Crow1.createCell((short)5).setCellValue(dt);
				   System.out.println("日均出勤时长（单位：小时）："+dt);
				   Crow1.createCell((short)6).setCellValue(late_15);
				   Crow1.createCell((short)7).setCellValue(late_30);
				   Crow1.createCell((short)8).setCellValue(late_60);
				   Crow1.createCell((short)9).setCellValue(leave_15);
				   Crow1.createCell((short)10).setCellValue(leave_30);
				   Crow1.createCell((short)11).setCellValue(leave_60);
				  
				   
		  	
		}
			  try {
					FileOutputStream fout = new FileOutputStream("D://Attendance//员工出勤统计表.xls");
					wb.write(fout);
					fout.close();
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} 
		}
	
	public static String[][] getData(HSSFSheet st,int ignoreRows) throws FileNotFoundException,IOException{
		List<String[]> result = new ArrayList<String[]>();
		int rowSize = 0;
		HSSFCell cell = null;
		for(int sheetIndex = 0; sheetIndex<st.getPhysicalNumberOfRows();sheetIndex++){
			for(int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum();rowIndex++){
				HSSFRow row = st.getRow(rowIndex);
				if(row == null){
					continue;
				}
				int tempRowSize = row.getLastCellNum()+1;
				if(tempRowSize>rowSize){
					rowSize = tempRowSize;
				}
				String[] values  = new String[rowSize];
				Arrays.fill(values, "");
				boolean hasValue = false;
				for(short columnIndex = 0; columnIndex <= row.getLastCellNum();columnIndex++){
					String value = "";
					cell = row.getCell(columnIndex);
					if(cell != null){
				
						switch(cell.getCellType()){
						case HSSFCell.CELL_TYPE_STRING:
							value = cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:
							if(HSSFDateUtil.isCellDateFormatted(cell)){
								Date date = cell.getDateCellValue();
								if(date != null){
									value = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
								}else{
									value = "";
								}
							}else{
								value = new DecimalFormat("0").format(cell.getNumericCellValue());
							}
							break;
						case HSSFCell.CELL_TYPE_FORMULA:
							break;
						case HSSFCell.CELL_TYPE_BLANK:
							break;
						case HSSFCell.CELL_TYPE_ERROR:
							break;
						case HSSFCell.CELL_TYPE_BOOLEAN:
							value = (cell.getBooleanCellValue() == true? "Y":"N");
							break;
						default:
							value ="";
						}
					}
					if(columnIndex == 0 && value.trim().equals("")){
						break;
					}
					values[columnIndex] = rightTrim(value);
					hasValue = true;
				}
				if(hasValue){
					result.add(values);
				}
			}
		}
		String[][] returnArray = new String[result.size()][rowSize];
		for(int i =0; i<returnArray.length;i++){
			returnArray[i] = result.get(i);
		}
		return returnArray;
	}
	public static String rightTrim(String str){
		if(str == null){
			return "";
		}
		int length = str.length();
		for(int i = length-1;i>=0;i--){
			if(str.charAt(i) != 0x20){
				break;
			}
			length--;
		}
		return str.substring(0,length);
	}
}