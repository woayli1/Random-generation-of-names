package main;


import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;


public class Main {

	public static void main(String[] args) {
		File file = new File("F:/excel/����.xlsx");
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
		System.out.println("��ʼ�����������");
		while(set.size()<=a){
			b=ExcelUtil.randomSet(b,u,m,(int) function(u,m), set, result);
			if(b>a){
				break;
			}
		}
		System.out.println("��ɣ�������"+set.size()+"��");
		for (Object s:set) {  
			result2.add(s); 
		} 
		System.out.println("��������excel");
		ExcelUtil.writeExcel(result2,"F:/excel/bb.xls");
		System.out.println("�����");
		System.out.println("�����ļ�λ��:F:/excel/bb.xls");
	}
	public static int function(int n,int m){
		return n*m+n*m*m;
	}
}
