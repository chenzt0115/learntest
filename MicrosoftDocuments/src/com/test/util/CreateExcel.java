package com.test.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	/**
     * ����һ��Excel
     * @param fileName �ļ���
     * @param dataList ����
     * @throws IOException
     */
    public static void buildXLSX(String fileName, List<Object[]> dataList) {
    	// ����һ��������
        XSSFWorkbook workBook = new XSSFWorkbook();
        FileOutputStream outStream = null;
        try
        {
            // ����һ������
            XSSFSheet sheet = workBook.createSheet();
            workBook.setSheetName(0,"info");
            
            XSSFCellStyle cellStyle = workBook.createCellStyle();
            XSSFDataFormat format = workBook.createDataFormat();
            cellStyle.setDataFormat(format.getFormat("@"));
            //�����赼��������
            for(int i=0;i<dataList.size();i++){
                XSSFRow row = sheet.createRow(i);
                Object[] oneRowData =  dataList.get(i);
                for(int j=0;j<oneRowData.length;j++)
                {
                	XSSFCell cell = row.createCell(j);
                	cell.setCellStyle(cellStyle);
                	Object value = oneRowData[j];
                	// �������ж����ݵ�����
					if (value instanceof String) {
						cell.setCellStyle(cellStyle);
						cell.setCellValue((String) value);
					}else if(value instanceof Integer){
    	        		cell.setCellValue((double)value);
    	        	}else if(value instanceof Double){
    	        		cell.setCellValue((double)value);
    	        	}else if(value instanceof BigDecimal){
    	        		cell.setCellValue((double)value);
    	        	}else if(value instanceof Boolean){
    	        		cell.setCellValue((boolean)value);
    	        	}
					
					int length = String.valueOf(value).getBytes().length;
					sheet.setColumnWidth(j, length * 256 + 150);
//					sheet.autoSizeColumn(j, true);
                }
            }
            File  file = new File(fileName);
            //�ļ������
            outStream = new FileOutputStream(file);
            workBook.write(outStream);
            outStream.flush();
            outStream.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        } finally {
        	try {
				outStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
        }
    }
    
    private static boolean isCnCharacter(String value){
    	return value.matches("[\u4e00-\u9fcc]+");
    }
    
    public static void main(String[] args) {
    	List<Object[]> dataList = new ArrayList<Object[]>();
    	String[] data = new String[]{"��0��","1.0","2.0","3.0","2018102711701662500001","����","55.16","û��\"��ʲô������\"","����,С��"};
    	dataList.add(data);
    	CreateExcel.buildXLSX("E:\\test\\test.xlsx", dataList);
	}
}