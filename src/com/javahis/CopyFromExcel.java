package com.javahis;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.tools.ant.Task;

public class CopyFromExcel extends Task {
	private Map<Integer, String> fileMap = new HashMap();
	/**
	 * 复制类型
	 */
	String copyType;
	/**
	 * 拷贝文件主目录
	 */
	String basedir;

	public String getCopyType() {
		return copyType;
	}

	public void setCopyType(String copyType) {
		this.copyType = copyType;
	}

	/**
	 * 拷贝文件目的目录
	 */
	String todir;
	/**
	 * excel 文件路径
	 */
	String excelFilePath;
	/**
	 * excel 文件中更新版本号(统一维护)字段
	 */
	String versionNumber;
	/**
	 * Class 文件路径前缀 如 com/javahis
	 */
	String classPrefixPath;
	/**
	 * 如果文件没有扩展名时自动追加该扩展名
	 */
	String expandedName;
	/**
	 * 初始读取行数
	 */
	int initRowNum;

	public int getInitRowNum() {
		return initRowNum;
	}

	public void setInitRowNum(int initRowNum) {
		this.initRowNum = initRowNum;
	}

	private CopyFromExcel.ExcelReader excelReader = null;

	public void execute() {
		System.out.println("拷贝分析……");
		if (basedir == null || basedir.equals("")) {
			System.err.println("属性 basedir 为空，无法确定源文件目录。");
			return;
		}
		if (todir == null || todir.equals("")) {
			System.err.println("属性 todir 为空，无法确定目的根目录。");
			return;
		}
		excelReader = new ExcelReader();
		try {
			fileMap = excelReader.readExcelColTOPath(getExcelFilePath());
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		copyFile();
	}

	/**
	 * 拷贝文件
	 */
	public void copyFile() {
		// System.out.println(i);
		if (fileMap == null || fileMap.isEmpty()) {
			return;
		}
		Iterator it = fileMap.entrySet().iterator();
		while (it.hasNext()) {
			Map.Entry<Integer, String> entry = (Map.Entry<Integer, String>) it.next();
			String[] fileValue = ((String) entry.getValue()).split("@");
			String placepath = fileValue[1];
			String f1 = this.basedir + "/" + fileValue[0];
			String f2 = this.todir + "/" + fileValue[0];
			try {
				boolean isCopyed = false;
				if ((placepath != null) && (placepath.trim() != "")) {
					if (placepath.indexOf("前端") != -1) {
						String qd = f2.replace("WEB-INF", "common");
						System.out.println("前端 拷贝 ：正在拷贝文件:" + f1 + "  到目标文件夹：" + qd);
						fileChannelCope(f1, qd);
						isCopyed = true;
					}
					if (f1.indexOf("/jdo/") >= 0) {
						f2 = f2.replaceFirst("classes", "jdoClass");
					}

					if (f1.indexOf("/config-Client/") >= 0) {
						f2 = f2.replaceFirst("config-Client/", "");
					}

					//
					if (f1.indexOf("/WebContent/config/") >= 0) {
						f2 = f2.replaceFirst("WebContent/config/", "");
					}

					if (f1.indexOf("/config-Server/") >= 0) {
						f2 = f2.replaceFirst("config/config-Server/", "WEB-INF/config/");
					}

					//
					if (f1.indexOf("/WebContent/WEB-INF/config/") >= 0) {
						f2 = f2.replaceFirst("config/WebContent/WEB-INF/config/", "WEB-INF/config/");
					}

					if (f1.indexOf("/action/") >= 0) {
						f2 = f2.replaceFirst("classes", "actionClass");
					}

					if (placepath.indexOf("后端") != -1) {
						System.out.println("后端 拷贝 ：正在拷贝文件:" + f1 + "  到目标文件夹：" + f2);
						fileChannelCope(f1, f2);
						isCopyed = true;
					}

				}

				if (!isCopyed) {
					fileChannelCope(f1, f2);
					System.out.println("拷贝 ：正在拷贝文件:" + f1 + "  到目标文件夹：" + f2);
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				System.err.println("拷贝文件出错。");
				e.printStackTrace();
			}
		}
	}

	/**
	 * excel 文件路径
	 */
	public String getExcelFilePath() {
		return excelFilePath;
	}

	/**
	 * Class 文件路径前缀 如 com/javahis
	 */
	public String getClassPrefixPath() {
		return classPrefixPath;
	}

	/**
	 * 如果文件没有扩展名时自动追加该扩展名
	 */
	public String getExpandedName() {
		return expandedName;
	}

	/**
	 * excel 文件路径
	 */
	public void setExcelFilePath(String excelFilePath) {
		this.excelFilePath = excelFilePath;
	}

	/**
	 * excel 文件中更新版本号(统一维护)字段
	 */
	public String getVersionNumber() {
		return versionNumber;
	}

	/**
	 * excel 文件中更新版本号(统一维护)字段
	 */
	public void setVersionNumber(String versionNumber) {
		this.versionNumber = versionNumber;
	}

	/**
	 * Class 文件路径前缀 如 com/javahis
	 */
	public void setClassPrefixPath(String classPrefixPath) {
		this.classPrefixPath = classPrefixPath;
	}

	/**
	 * 如果文件没有扩展名时自动追加该扩展名
	 */
	public void setExpandedName(String expandedName) {
		this.expandedName = expandedName;
	}

	public String getBasedir() {
		return basedir;
	}

	public void setBasedir(String basedir) {
		this.basedir = basedir;
	}

	public String getTodir() {
		return todir;
	}

	public void setTodir(String todir) {
		this.todir = todir;
	}

	/**
	 * 文件拷贝
	 * 
	 * @param f1
	 * @param f2
	 * @return
	 * @throws Exception
	 */
	public static long fileChannelCope(String filePath1, String filePath2) throws Exception {
		File f1 = new File(filePath1);
		File f2 = new File(filePath2);
		File fdir = new File(filePath2.substring(0, filePath2.lastIndexOf("/")));
		fdir.mkdirs();
		long time = new Date().getTime();
		int length = 2097152;
		FileInputStream in = new FileInputStream(f1);
		FileOutputStream out = new FileOutputStream(f2);
		FileChannel inC = in.getChannel();
		FileChannel outC = out.getChannel();
		ByteBuffer b = null;
		while (true) {
			if (inC.position() == inC.size()) {
				inC.close();
				outC.close();
				return new Date().getTime() - time;
			}
			if ((inC.size() - inC.position()) < length) {
				length = (int) (inC.size() - inC.position());
			} else
				length = 2097152;
			b = ByteBuffer.allocateDirect(length);
			inC.read(b);
			b.flip();
			outC.write(b);
			outC.force(false);
		}
	}

	public static void main(String[] args) {
		// File f = new File("F:/安装/sjsk2.avi");
		// File f1 = new File("c:/sjsk2.avi");
		// try {
		// CopyFromExcel.fileChannelCope(f, f1);
		// } catch (Exception e) {
		// // TODO Auto-generated catch block
		// e.printStackTrace();
		// }
	}

	/**
	 * 操作Excel表格的功能类
	 * 
	 * @author：xueyf
	 * @version 1.0
	 */
	class ExcelReader {
		private POIFSFileSystem fs;
		private HSSFWorkbook wb;
		private HSSFSheet sheet;
		private HSSFRow row;

		/**
		 * 读取Excel表格表头的内容
		 * 
		 * @param InputStream
		 * @return String 表头内容的数组
		 * 
		 */
		public String[] readExcelTitle(InputStream is) {
			try {
				fs = new POIFSFileSystem(is);
				wb = new HSSFWorkbook(fs);
			} catch (IOException e) {
				e.printStackTrace();
			}
			sheet = wb.getSheetAt(0);
			row = sheet.getRow(0);
			// 标题总列数
			int colNum = row.getPhysicalNumberOfCells();
			String[] title = new String[colNum];
			for (int i = 0; i < colNum; i++) {
				title[i] = getStringCellValue(row.getCell((short) i));
			}
			return title;
		}

		/**
		 * 读取Excel数据内容
		 * 
		 * @param InputStream
		 * @return Map 包含单元格数据内容的Map对象
		 * @throws FileNotFoundException
		 */
		public Map<Integer, String> readExcelColTOPath(String filePath) throws FileNotFoundException {
			InputStream is = new FileInputStream(getExcelFilePath());
			Map<Integer, String> content = new HashMap<Integer, String>();
			String str = "";
			try {
				fs = new POIFSFileSystem(is);
				wb = new HSSFWorkbook(fs);
			} catch (IOException e) {
				e.printStackTrace();
			}
			sheet = wb.getSheetAt(0);
			// 得到总行数
			int rowNum = sheet.getLastRowNum();
			row = sheet.getRow(0);
			// 正文内容应该从第二行开始,第一行为表头的标题
			int count = 0;
			int i = getInitRowNum() > 1 ? getInitRowNum() - 1 : 1;
			for (; i <= rowNum; i++) {
				str = "";
				row = sheet.getRow(i);
				int j = 0;
				try {
					str += getStringCellValue(row.getCell(3)).trim() + "/" + getStringCellValue(row.getCell(4)).trim();
					if (str.length() == 1 && str.equals("/")) {
						break;
					}

					// 追加文件前缀路径
					str = getClassPrefixPath() == null || getClassPrefixPath().equals("") ? str
							: getClassPrefixPath() + str;
					// 斜杠到反斜杠的替换
					str = str.replaceAll("\\\\", "\\/");
					// 追加文件扩展名
					str = str.lastIndexOf(".") == -1 ? str + (getExpandedName() == null ? "" : getExpandedName()) : str;
					// 文件扩展名比较，若与当前拷贝扩展名不同，不予拷贝
					if ((str.split("\\.").length == 1)
							|| (getCopyType() != null && !str.substring(str.lastIndexOf(".")).equals(getCopyType()))) {
						continue;
					}
				} catch (Exception ex) {
					System.out.println("共查询到" + i + "个更新文件");
					ex.printStackTrace();
					break;
				}
				j++;
				count++;
				str = str + "@" + getStringCellValue(this.row.getCell(2)) + " ";
				content.put(i, str);

			}
			System.out.println("共查询到" + count + "个更新文件");
			return content;
		}

		/**
		 * 读取Excel数据内容
		 * 
		 * @param InputStream
		 * @return Map 包含单元格数据内容的Map对象
		 */
		public Map<Integer, String> readExcelContent(InputStream is) {
			Map<Integer, String> content = new HashMap<Integer, String>();
			String str = "";
			try {
				fs = new POIFSFileSystem(is);
				wb = new HSSFWorkbook(fs);
			} catch (IOException e) {
				e.printStackTrace();
			}
			sheet = wb.getSheetAt(0);
			// 得到总行数
			int rowNum = sheet.getLastRowNum();
			row = sheet.getRow(0);
			int colNum = row.getPhysicalNumberOfCells();
			// 正文内容应该从第二行开始,第一行为表头的标题
			for (int i = 1; i <= rowNum; i++) {
				row = sheet.getRow(i);
				int j = 0;
				while (j < colNum) {
					// 每个单元格的数据内容用"-"分割开，以后需要时用String类的replace()方法还原数据
					// 也可以将每个单元格的数据设置到一个javabean的属性中，此时需要新建一个javabean
					str += getStringCellValue(row.getCell((short) j)).trim() + "-";
					j++;
				}
				content.put(i, str);
				str = "";
			}
			return content;
		}

		/**
		 * 获取单元格数据内容为字符串类型的数据
		 * 
		 * @param cell Excel单元格
		 * @return String 单元格数据内容
		 */
		private String getStringCellValue(HSSFCell cell) {
			String strCell = "";
			switch (cell.getCellType()) {
			case HSSFCell.CELL_TYPE_STRING:
				strCell = cell.getStringCellValue();
				break;
			case HSSFCell.CELL_TYPE_NUMERIC:
				strCell = String.valueOf(cell.getNumericCellValue());
				break;
			case HSSFCell.CELL_TYPE_BOOLEAN:
				strCell = String.valueOf(cell.getBooleanCellValue());
				break;
			case HSSFCell.CELL_TYPE_BLANK:
				strCell = "";
				break;
			default:
				strCell = "";
				break;
			}
			if (strCell.equals("") || strCell == null) {
				return "";
			}
			if (cell == null) {
				return "";
			}
			return strCell;
		}

		/**
		 * 获取单元格数据内容为日期类型的数据
		 * 
		 * @param cell Excel单元格
		 * @return String 单元格数据内容
		 */
		private String getDateCellValue(HSSFCell cell) {
			String result = "";
			try {
				int cellType = cell.getCellType();
				if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
					Date date = cell.getDateCellValue();
					result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1) + "-" + date.getDate();
				} else if (cellType == HSSFCell.CELL_TYPE_STRING) {
					String date = getStringCellValue(cell);
					result = date.replaceAll("[年月]", "-").replace("日", "").trim();
				} else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
					result = "";
				}
			} catch (Exception e) {
				System.out.println("日期格式不正确!");
				e.printStackTrace();
			}
			return result;
		}

	}

}
