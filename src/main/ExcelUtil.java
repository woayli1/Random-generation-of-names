package main;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashSet;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	//Ĭ�ϵ�Ԫ������Ϊ����ʱ��ʽ
	private static DecimalFormat df = new DecimalFormat("0");
	// Ĭ�ϵ�Ԫ���ʽ�������ַ��� 
	private static SimpleDateFormat sdf = new SimpleDateFormat(  "yyyy-MM-dd HH:mm:ss"); 
	// ��ʽ������
	private static DecimalFormat nf = new DecimalFormat("0.00");  
	public static ArrayList<ArrayList<Object>> readExcel(File file){
		if(file == null){
			return null;
		}
		if(file.getName().endsWith("xlsx")){
			//����ecxel2007
			return readExcel2007(file);
		}else{
			//����ecxel2003
			return readExcel2003(file);
		}
	}
	/*
	 * @return �����ؽ���洢��ArrayList�ڣ��洢�ṹ���λ��������
	 * lists.get(0).get(0)��ʾ��ȥExcel��0��0�е�Ԫ��
	 */
	public static ArrayList<ArrayList<Object>> readExcel2003(File file){
		try{
			ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();
			ArrayList<Object> colList;
			HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
			HSSFSheet sheet = wb.getSheetAt(0);
			HSSFRow row;
			HSSFCell cell;
			Object value;
			for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
				row = sheet.getRow(i);
				colList = new ArrayList<Object>();
				if(row == null){
					//����ȡ��Ϊ��ʱ
					if(i != sheet.getPhysicalNumberOfRows()){//�ж��Ƿ������һ��
						rowList.add(colList);
					}
					continue;
				}else{
					rowCount++;
				}
				for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
					cell = row.getCell(j);
					if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
						//���õ�Ԫ��Ϊ��
						if(j != row.getLastCellNum()){//�ж��Ƿ��Ǹ��������һ����Ԫ��
							colList.add("");
						}
						continue;
					}
					switch(cell.getCellType()){
					case XSSFCell.CELL_TYPE_STRING:  
						System.out.println(i + "��" + j + " �� is String type");  
						value = cell.getStringCellValue();  
						break;  
					case XSSFCell.CELL_TYPE_NUMERIC:  
						if ("@".equals(cell.getCellStyle().getDataFormatString())) {  
							value = df.format(cell.getNumericCellValue());  
						} else if ("General".equals(cell.getCellStyle()  
								.getDataFormatString())) {  
							value = nf.format(cell.getNumericCellValue());  
						} else {  
							value = sdf.format(HSSFDateUtil.getJavaDate(cell  
									.getNumericCellValue()));  
						}  
						System.out.println(i + "��" + j  
								+ " �� is Number type ; DateFormt:"  
								+ value.toString()); 
						break;  
					case XSSFCell.CELL_TYPE_BOOLEAN:  
						System.out.println(i + "��" + j + " �� is Boolean type");  
						value = Boolean.valueOf(cell.getBooleanCellValue());
						break;  
					case XSSFCell.CELL_TYPE_BLANK:  
						System.out.println(i + "��" + j + " �� is Blank type");  
						value = "";  
						break;  
					default:  
						System.out.println(i + "��" + j + " �� is default type");  
						value = cell.toString();  
					}// end switch
					colList.add(value);
				}//end for j
				rowList.add(colList);
			}//end for i

			return rowList;
		}catch(Exception e){
			return null;
		}
	}

	public static ArrayList<ArrayList<Object>> readExcel2007(File file){
		try{
			ArrayList<ArrayList<Object>> rowList = new ArrayList<ArrayList<Object>>();
			ArrayList<Object> colList;
			XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
			XSSFSheet sheet = wb.getSheetAt(0);
			XSSFRow row;
			XSSFCell cell;
			Object value;
			for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
				row = sheet.getRow(i);
				colList = new ArrayList<Object>();
				if(row == null){
					//����ȡ��Ϊ��ʱ
					if(i != sheet.getPhysicalNumberOfRows()){//�ж��Ƿ������һ��
						rowList.add(colList);
					}
					continue;
				}else{
					rowCount++;
				}
				for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
					cell = row.getCell(j);
					if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
						//���õ�Ԫ��Ϊ��
						if(j != row.getLastCellNum()){//�ж��Ƿ��Ǹ��������һ����Ԫ��
							colList.add("");
						}
						continue;
					}
					switch(cell.getCellType()){
					case XSSFCell.CELL_TYPE_STRING:  
						//System.out.println(i + "��" + j + " �� is String type");  
						value = cell.getStringCellValue();  
						break;  
					case XSSFCell.CELL_TYPE_NUMERIC:  
						if ("@".equals(cell.getCellStyle().getDataFormatString())) {  
							value = df.format(cell.getNumericCellValue());  
						} else if ("General".equals(cell.getCellStyle()  
								.getDataFormatString())) {  
							value = nf.format(cell.getNumericCellValue());  
						} else {  
							value = sdf.format(HSSFDateUtil.getJavaDate(cell  
									.getNumericCellValue()));  
						}  
						System.out.println(i + "��" + j  
								+ " �� is Number type ; DateFormt:"  
								+ value.toString()); 
						break;  
					case XSSFCell.CELL_TYPE_BOOLEAN:  
						System.out.println(i + "��" + j + " �� is Boolean type");  
						value = Boolean.valueOf(cell.getBooleanCellValue());
						break;  
					case XSSFCell.CELL_TYPE_BLANK:  
						System.out.println(i + "��" + j + " �� is Blank type");  
						value = "";  
						break;  
					default:  
						System.out.println(i + "��" + j + " �� is default type");  
						value = cell.toString();  
					}// end switch
					colList.add(value);
				}//end for j
				rowList.add(colList);
			}//end for i

			return rowList;
		}catch(Exception e){
			System.out.println("exception");
			return null;
		}
	}

	public static int randomSet(int b,int max, int ma2, int n,HashSet<Object> set, ArrayList<ArrayList<Object>> result) {  
		String names; //��
		String pwd; //��
		String pwd2; //��2
		String np;
		int k,j,j2;
		for (int i = 1; i <= max; i++) {  
			// ����Math.random()����  
			 k = (int) (Math.random() * (max - 1)) + 1;
			 j = (int) (Math.random() * (ma2 - 1)) + 1;
			 j2 = (int) (Math.random() * (ma2 - 1)) + 1;
			do{
			names=result.get(k).get(0).toString();
			pwd=result.get(j).get(1).toString();
			pwd2=result.get(j2).get(1).toString();
			k = (int) (Math.random() * (max - 1)) + 1;
			 j = (int) (Math.random() * (ma2 - 1)) + 1;
			 j2 = (int) (Math.random() * (ma2 - 1)) + 1;}
			while(names.equals(null) || pwd.equals(null) || pwd2.equals(null));

			while(names.length()==0 || pwd.length()==0 || pwd2.length()==0)
			{
				names=result.get(k).get(0).toString().trim();
				pwd=result.get(j).get(1).toString().trim();
				pwd2=result.get(j2).get(1).toString().trim();
				k = (int) (Math.random() * (max - 1)) + 1;
				 j = (int) (Math.random() * (ma2 - 1)) + 1;
				 j2 = (int) (Math.random() * (ma2 - 1)) + 1;
			}
			np=names+pwd+pwd2;
			if(!set.add(np)){
				b++;
			};
		}  
		int setSize = set.size();
		// ����������С��ָ�����ɵĸ���������õݹ�������ʣ�����������������ѭ����ֱ���ﵽָ����С  
		if (setSize < n) {  
			randomSet(b,max, ma2, n - setSize,set,result);// �ݹ�  
		}  
		return b;
	}  

	private String np;
	@Override  
	public boolean equals(Object obj) {  
		if (this == obj)  
			return true;  
		if (obj == null)  
			return false;  
		if (getClass() != obj.getClass())  
			return false;  
		ExcelUtil other = (ExcelUtil) obj;  

		if (np != other.np)  
			return false;  
		else if (!np.equals(other.np))  
			return false;  
		return true;  
	}  

	public static void writeExcel(ArrayList<Object> result,String path){
		if(result == null){
			return;
		}
		HSSFCell cell;
		int i=0,k=0;
		HSSFWorkbook wb = new HSSFWorkbook();
		for(int a=1;a<=100;a++){
		HSSFSheet sheet = wb.createSheet("sheet"+a);
		for(k=0;i < result.size(); i++,k++){
			if(k==65536)break;
			HSSFRow row = sheet.createRow(k);
			if(result.get(i) != null){
				for(int j = 0; j < 1; j ++){
					cell = row.createCell(j);
					cell.setCellValue(result.get(i).toString());
				}

			}
		}
		}
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try
		{
			System.out.println("���ڽ��������д��excel");
			wb.write(os);
		} catch (IOException e){
			e.printStackTrace();
		}
		byte[] content = os.toByteArray();
		File file = new File(path);//Excel�ļ����ɺ�洢��λ�á�
		OutputStream fos  = null;
		try
		{
			fos = new FileOutputStream(file);
			fos.write(content);
			os.close();
			fos.close();
		}catch (Exception e){
			e.printStackTrace();
		}
		return ;
	}

	public static DecimalFormat getDf() {
		return df;
	}
	public static void setDf(DecimalFormat df) {
		ExcelUtil.df = df;
	}
	public static SimpleDateFormat getSdf() {
		return sdf;
	}
	public static void setSdf(SimpleDateFormat sdf) {
		ExcelUtil.sdf = sdf;
	}
	public static DecimalFormat getNf() {
		return nf;
	}
	public static void setNf(DecimalFormat nf) {
		ExcelUtil.nf = nf;
	}



}
