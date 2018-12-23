package com.test.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ModifyExcel {
	/** ������Ϣ */
	private String errorInfo;

	/**
	 * @��������֤excel�ļ�
	 * @���ߣ�
	 * @ʱ�䣺2012-08-29 ����16:27:15
	 * @������@param filePath���ļ�����·��
	 * @������@return
	 * @����ֵ��boolean
	 */
	public boolean validateExcel(String filePath) {
		/** ����ļ����Ƿ�Ϊ�ջ����Ƿ���Excel��ʽ���ļ� */

		if (filePath == null
				|| !(WDWUtil.isExcel2003(filePath) || WDWUtil
						.isExcel2007(filePath))) {
			errorInfo = "�ļ�������excel��ʽ";
			return false;
		}

		/** ����ļ��Ƿ���� */
		File file = new File(filePath);
		if (file == null || !file.exists()) {
			errorInfo = "�ļ�������";
			return false;
		}
		return true;
	}

	/**
	 * 
	 * @param fileName
	 * @param dataList
	 */
	public void batchModifyXLSX(String fileName, List<Object[]> dataList) {
		try {
			/** ��֤�ļ��Ƿ�Ϸ� */
			if (!validateExcel(fileName)) {
				System.out.println(errorInfo);
				return;
			}

			/** �ж��ļ������ͣ���2003����2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(fileName)) {
				isExcel2003 = false;
			}

			batchModify(fileName, isExcel2003, dataList);
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
	public void modifyXLSX(String fileName, int row, int col, Object value) {
		try {
			/** ��֤�ļ��Ƿ�Ϸ� */
			if (!validateExcel(fileName)) {
				System.out.println(errorInfo);
				return;
			}

			/** �ж��ļ������ͣ���2003����2007 */
			boolean isExcel2003 = true;
			if (WDWUtil.isExcel2007(fileName)) {
				isExcel2003 = false;
			}

			modify(fileName, isExcel2003, row, col, value);
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private void modify(String fileName, boolean isExcel2003, int row, int col,
			Object value) {
		InputStream is = null;
		FileOutputStream outStream = null;
		try {
			/** ���ñ����ṩ�ĸ�������ȡ�ķ��� */
			File file = new File(fileName);
			is = new FileInputStream(file);

			/** ���ݰ汾ѡ�񴴽�Workbook�ķ�ʽ */
			Workbook wb = null;
			if (isExcel2003) {
				wb = new HSSFWorkbook(is);
			} else {
				wb = new XSSFWorkbook(is);
			}
			modify(wb, row, col, value);
			is.close();

			// �ļ������
			outStream = new FileOutputStream(file);
			wb.write(outStream);
			outStream.flush();
			outStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					is = null;
					e.printStackTrace();
				}
			}
			if (outStream != null) {
				try {
					outStream.close();
				} catch (IOException e) {
					outStream = null;
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * @��������ȡ����
	 * @���ߣ�
	 * @ʱ�䣺2012-08-29 ����16:50:15
	 * @������@param Workbook
	 * @������@return
	 * @����ֵ��List<List<String>>
	 */
	public void modify(Workbook wb, int row, int col, Object value) {
		/** �õ���һ��shell */
		Sheet sheet = wb.getSheetAt(0);
		CellStyle cellStyle = wb.createCellStyle();
		DataFormat format = wb.createDataFormat();
		cellStyle.setDataFormat(format.getFormat("@"));

		Cell cell = sheet.getRow(row).getCell(col);
		if (null != cell) {
			// �������ж����ݵ�����
			if (value instanceof String) {
				cell.setCellStyle(cellStyle);
				cell.setCellValue((String) value);
			} else if (value instanceof Integer) {
				cell.setCellValue((double) value);
			} else if (value instanceof Double) {
				cell.setCellValue((double) value);
			} else if (value instanceof BigDecimal) {
				cell.setCellValue((double) value);
			} else if (value instanceof Boolean) {
				cell.setCellValue((boolean) value);
			}
		}
		int length = String.valueOf(value).getBytes().length;
		sheet.setColumnWidth(col, length * 256 + 150);
//		sheet.autoSizeColumn(col, true);
	}
	
	private void batchModify(String fileName, boolean isExcel2003, List<Object[]> dataList) {
		InputStream is = null;
		FileOutputStream outStream = null;
		try {
			/** ���ñ����ṩ�ĸ�������ȡ�ķ��� */
			File file = new File(fileName);
			is = new FileInputStream(file);

			/** ���ݰ汾ѡ�񴴽�Workbook�ķ�ʽ */
			Workbook wb = null;
			if (isExcel2003) {
				wb = new HSSFWorkbook(is);
			} else {
				wb = new XSSFWorkbook(is);
			}
			batchModify(wb, dataList);
			is.close();

			// �ļ������
			outStream = new FileOutputStream(file);
			wb.write(outStream);
			outStream.flush();
			outStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					is = null;
					e.printStackTrace();
				}
			}
			if (outStream != null) {
				try {
					outStream.close();
				} catch (IOException e) {
					outStream = null;
					e.printStackTrace();
				}
			}
		}
	}
	
	private void batchModify(Workbook wb, List<Object[]> dataList) {
		/** �õ���һ��shell */
		Sheet sheet = wb.getSheetAt(0);
		CellStyle cellStyle = wb.createCellStyle();
		DataFormat format = wb.createDataFormat();
		cellStyle.setDataFormat(format.getFormat("@"));

		for (int i = 0; i < dataList.size(); i++) {
			Object[] rowdata = dataList.get(i);
			Row row = sheet.getRow(i);
			if(null != row){
				for (int j = 0; j < rowdata.length; j++) {
					Cell cell = row.getCell(j);
					if(null != cell && null != rowdata[j]){
						Object value = rowdata[j];
						// �������ж����ݵ�����
						if (value instanceof String) {
							cell.setCellStyle(cellStyle);
							cell.setCellValue((String) value);
						} else if (value instanceof Integer) {
							cell.setCellValue((double) value);
						} else if (value instanceof Double) {
							cell.setCellValue((double) value);
						} else if (value instanceof BigDecimal) {
							cell.setCellValue((double) value);
						} else if (value instanceof Boolean) {
							cell.setCellValue((boolean) value);
						}
					}
				}
			} else {
				break;
			}
		}
	}

	private boolean isCnCharacter(String value){
    	return value.matches("[\u4e00-\u9fcc]+");
    }
	
	public static void main(String[] args) {
		ModifyExcel poi = new ModifyExcel();
		poi.modifyXLSX("E:\\test\\test.xlsx", 0, 5, "������");
	}
}