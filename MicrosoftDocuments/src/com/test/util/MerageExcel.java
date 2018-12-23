package com.test.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MerageExcel {
    private Map<String, Integer> ywlshIndex = null;
    
	public void merage(String fname1, String fname2, String destname){
		writeone(fname1, destname);
		writetwo(fname2, destname);
	}
	
	private void writetwo(String fname2, String destname){
		// 声明一个工作薄
		InputStream is = null;
    	FileOutputStream outStream = null;
        try
        {
        	//读取源文件内容
            ImportExecl importer = new ImportExecl();
            List<List<String>> res = importer.read(fname2);
    		if(null == res || res.size() == 0){
    			return;
    		}
    		
    		
    		 /** 调用本类提供的根据流读取的方法 */
            File file = new File(destname);
            is = new FileInputStream(file);
            
            /** 根据版本选择创建Workbook的方式 */
            Workbook  wb = new XSSFWorkbook(is);
            ModifyExcel modifyer = new ModifyExcel();
            
            /** 循环Excel的行 */
            for (int r = 0; r < res.size(); r++)
            {
            	List<String> row = res.get(r);
                
                String ywlsh = row.get(0);
                if(!ywlshIndex.containsKey(ywlsh)){
                	continue;
                }
                int idx = ywlshIndex.remove(ywlsh);
                
                String lcdm = row.get(1);
                modifyer.modify(wb, idx, 1, lcdm);
            }
            
            is.close();
            
            //文件输出流
            outStream = new FileOutputStream(file);
            wb.write(outStream);
            outStream.flush();
            outStream.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }finally
        {
            if (is != null)
            {
                try
                {
                    is.close();
                }
                catch (IOException e)
                {
                    is = null;
                    e.printStackTrace();
                }
            }
            if (outStream != null)
            {
                try
                {
                	outStream.close();
                }
                catch (IOException e)
                {
                	outStream = null;
                    e.printStackTrace();
                }
            }
        }
	}
	
	private void writeone(String fname1, String destname){
        try {
        	//读取源文件内容
            ImportExecl importer = new ImportExecl();
            List<List<String>> res = importer.read(fname1);
    		if(null == res || res.size() == 0){
    			return;
    		}
    		
    		ywlshIndex = new HashMap<String, Integer>();
    		List<Object[]> dataList = new ArrayList<Object[]>();
            for (int r = 0; r < res.size(); r++)
            {
            	List<String> row = res.get(r);
                
            	String ywlsh = row.get(3);
            	ywlshIndex.put(ywlsh, r);
            	
            	String bwnr = row.get(4);
            	String[] bwnrs = bwnr.split(",");
            	
				String[] data = new String[] { ywlsh, "lcdm", bwnrs[0],
						bwnrs[1], row.get(6) };
            	dataList.add(data);
            }
            CreateExcel.buildXLSX(destname, dataList);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		MerageExcel merge  = new MerageExcel();
		merge.merage("E:\\test\\运单.xlsx", "E:\\test\\test.xlsx", "E:\\test\\新文件.xlsx");
	}
}
