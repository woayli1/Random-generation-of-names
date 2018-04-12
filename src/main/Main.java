package main;


import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;


public class Main {

	public static void main(String[] args) {
		File file = new File("F:/excel/姓名.xlsx");
		HashSet<Object> set=new HashSet<Object>();
		int u=0;
		int m=0;
		ArrayList<ArrayList<Object>> result = ExcelUtil.readExcel(file);
		ArrayList<Object> result2 =new ArrayList<Object>();
		for(int i = 0 ;i < result.size() ;i++){
			for(int j = 0;j<result.get(i).size(); j++){
				if(j==0){
					if(result.get(i).get(j).toString().trim().length()>0)
						u++;}
				else{
					if(result.get(i).get(j).toString().trim().length()>0)
						m++;}
			}
		}
		int b=0;
		int a=4000000;
		System.out.println("开始生成随机姓名");
		while(set.size()<=a){
			b=ExcelUtil.randomSet(b,u,m,(int) function(u,m), set, result);
			if(b>a){
				break;
			}
		}
		System.out.println("完成，共生成"+set.size()+"个");
		for (Object s:set) {  
			result2.add(s); 
		} 
		System.out.println("正在生成excel");
		ExcelUtil.writeExcel(result2,"F:/excel/bb.xls");
		System.out.println("已完成");
		System.out.println("生成文件位置:F:/excel/bb.xls");
	}
	public static int function(int n,int m){
		return n*m+n*m*m;
	}
}
