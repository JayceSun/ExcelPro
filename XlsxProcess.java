package com.bonc.excelxlsx;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;

import com.bonc.customer.Utils;

public class XlsxProcess {

	
	public static void main(String[] args) {
		
		 System.out.println("=============================================");
		 
		 String inputDir = "E:\\gowork\\pholcus_pkg\\text_out";
		 String outputDir = "C:\\Users\\Administrator\\Desktop\\运行\\output";
		 
		 outputfile = outputDir + File.separator + outputFileName;
		 
		 processDir(inputDir,outputDir);
		 
		 FileOutputStream fileOut = null;
	        try {
	            fileOut = new FileOutputStream(outputfile);
	            wbCreat.write(fileOut);
	            wbCreat.close();
	            fileOut.close();
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	}
	
	/**
	 * 需要处理的文件格式
	 */
	private static String fileRm1 = ".xlsx";
	
	private static String fileRm2 = ".xlsx";
	
	//文件表示ID
    private static int fileID = 0;
    
    // 创建新的excel
    static XSSFWorkbook wbCreat = new XSSFWorkbook();
    
    static Sheet sheetCreat = wbCreat.createSheet();
    
    static String outputfile = null;
    
    static String outputFileName = "output.xlsx";
    
    static int destI = 0;
	
	/**
	 * 读取文件夹里面所有.xlsx 或者 .xls 文件
	 * @param inputDir
	 * @param outputDir
	 */
	public static void processDir(String inputDir, String outputDir) {
		 
		if (inputDir.endsWith("\\") || inputDir.endsWith(File.separator)) {
			inputDir = inputDir.substring(0, inputDir.length() - 1);
		}
		if (outputDir.endsWith("\\") || outputDir.endsWith(File.separator)) {
			outputDir = outputDir.substring(0, outputDir.length() - 1);
		}
		File dirFile = new File(inputDir);
		File outputFile = new File(outputDir);
		if (!outputFile.exists()) {
			outputFile.mkdirs();
		}
		if (dirFile != null) {
			File[] texts = dirFile.listFiles(new FileFilter() {
				// file 过滤目录文件名
				@Override
				public boolean accept(File pathname) {
					String lowerCase = pathname.getName().toLowerCase();
					return (lowerCase.endsWith(fileRm1) || lowerCase.endsWith(fileRm2))
							&& pathname.canRead();
				}
			});
			for (int i = 0; i < texts.length; i++) {
				String absPath = texts[i].getAbsolutePath();
//				outputfile = outputDir + File.separator
//						+ texts[i].getName();
				System.out.println(absPath);
				try {
					 processBySheet(absPath, outputfile);
					 fileID++;
				} catch (Exception e) {
					e.printStackTrace();
				}
			}

			File[] dirs = dirFile.listFiles(new FileFilter() {
				// file 过滤目录文件名
				@Override
				public boolean accept(File pathname) {
					return pathname.isDirectory();
				}
			});
			for (File dir : dirs) {
				String fileNewDir = dir.getAbsolutePath();
				String outputNewDir = fileNewDir.replace(inputDir, outputDir);
				processDir(fileNewDir, outputNewDir);
			}
		}
	}
	
	/**
	 * 读取并创建
	 * @param srcExcelName
	 * @param destPrefixName
	 */
	public static void processBySheet(String srcExcelName, String destPrefixName) {

		// 读取
		Workbook wb = readExcel(srcExcelName);
		// 创建新的excel
//		XSSFWorkbook wbCreat = new XSSFWorkbook();
		System.out.println(wb.getNumberOfSheets());
		
		for (int i = 0; i < wb.getNumberOfSheets(); i++) {
			try {
				Sheet sheet = wb.getSheetAt(i);
//				Sheet sheetCreat = wbCreat.createSheet(sheet.getSheetName());

				copySheetFilterAddCategory(sheet, sheetCreat, true,i);// 添加类别并过滤
			} catch (Exception e) {
				e.printStackTrace();
			}finally {
	            if (wb != null) {
	                try {
	                    wb.close();
	                } catch (IOException e) {
	                    e.printStackTrace();
	                }
	            }
	        }

		}
	}
	
	
	/**
	 * 读取excel
	 * @param filePath
	 * @return
	 */
	public static Workbook readExcel(String filePath) {
		Workbook wb = null;
		if (filePath == null) {
			return null;
		}
		String extString = filePath.substring(filePath.lastIndexOf("."));
		InputStream is = null;
		try {
			is = new FileInputStream(filePath);
			if (".xls".equals(extString)) {
				return wb = new HSSFWorkbook(is);
			} else if (".xlsx".equals(extString)) {
				return wb = new XSSFWorkbook(is);
			} else {
				System.out.println("不存在 xls && xlsx 文件！！");
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wb;
	}

	
	/**
	 * copy的方法
	 * @param srcSheet
	 * @param destSheet
	 * @param bIsFilter
	 * @return
	 */
	public static boolean copySheetFilterAddCategory(Sheet srcSheet, Sheet destSheet, boolean bIsFilter,int j) {
    	try {
            int firstRow = srcSheet.getFirstRowNum();
            int lastRow = srcSheet.getLastRowNum();
            
            for (int i = firstRow; i <= lastRow; i++) {
                // 取得源有excel Sheet的行
                Row srcRow = srcSheet.getRow(i);
                
                Row destRow = null;
                if (i == firstRow) {
                	if (j != 0) {
                        continue;
                    }

                    if (fileID != 0) {
                        continue;
                    }
                    // 创建新建excel Sheet的行
                    destRow = destSheet.createRow(destI++);
                    Utils.copyRowAddCell(srcRow, destRow);
                    continue;
                }

                if (srcRow == null) {
                    continue;
                }

                if (destI >= 60000) {
                    sheetCreat = wbCreat.createSheet(System.currentTimeMillis() + "");
                    destI = 0;
                }
                
                destRow = destSheet.createRow(destI++);

//                System.out.println(srcRow.getCell(7));
//                String sntmnt = Category.getSntmnt(srcRow);
                
                Utils.copyRowAddCell(srcRow, destRow);
            }
            return true;
        } catch (Exception e) {
        	e.printStackTrace();
            return false;
        }

    }
}
