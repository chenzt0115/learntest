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
     * 创建一个Excel
     * @param fileName 文件名
     * @param dataList 数据
     * @throws IOException
     */
    public static void buildXLSX(String fileName, List<Object[]> dataList) {
    	// 声明一个工作薄
        XSSFWorkbook workBook = new XSSFWorkbook();
        FileOutputStream outStream = null;
        try
        {

            // 生成一个表格
            XSSFSheet sheet = workBook.createSheet();
            workBook.setSheetName(0,"info");
            
            XSSFCellStyle cellStyle = workBook.createCellStyle();
            XSSFDataFormat format = workBook.createDataFormat();
            cellStyle.setDataFormat(format.getFormat("@"));
            //插入需导出的数据
            for(int i=0;i<dataList.size();i++){
                XSSFRow row = sheet.createRow(i);
                Object[] oneRowData =  dataList.get(i);
                for(int j=0;j<oneRowData.length;j++)
                {
                	XSSFCell cell = row.createCell(j);
                	cell.setCellStyle(cellStyle);
                	Object value = oneRowData[j];
                	// 以下是判断数据的类型
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
            //文件输出流
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
    	String[] data = new String[]{"第0行","1.0","2.0","3.0","2018102711701662500001","林增","55.16","没有\"事什么不可能\"","大于,小于"};
    	dataList.add(data);
    	CreateExcel.buildXLSX("E:\\test\\test.xlsx", dataList);
	}
}
