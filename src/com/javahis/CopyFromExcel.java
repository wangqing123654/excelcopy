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
	 * ��������
	 */
	String copyType;
	/**
	 * �����ļ���Ŀ¼
	 */
	String basedir;

	public String getCopyType() {
		return copyType;
	}

	public void setCopyType(String copyType) {
		this.copyType = copyType;
	}

	/**
	 * �����ļ�Ŀ��Ŀ¼
	 */
	String todir;
	/**
	 * excel �ļ�·��
	 */
	String excelFilePath;
	/**
	 * excel �ļ��и��°汾��(ͳһά��)�ֶ�
	 */
	String versionNumber;
	/**
	 * Class �ļ�·��ǰ׺ �� com/javahis
	 */
	String classPrefixPath;
	/**
	 * ����ļ�û����չ��ʱ�Զ�׷�Ӹ���չ��
	 */
	String expandedName;
	/**
	 * ��ʼ��ȡ����
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
		System.out.println("������������");
		if (basedir == null || basedir.equals("")) {
			System.err.println("���� basedir Ϊ�գ��޷�ȷ��Դ�ļ�Ŀ¼��");
			return;
		}
		if (todir == null || todir.equals("")) {
			System.err.println("���� todir Ϊ�գ��޷�ȷ��Ŀ�ĸ�Ŀ¼��");
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
	 * �����ļ�
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
					if (placepath.indexOf("ǰ��") != -1) {
						String qd = f2.replace("WEB-INF", "common");
						System.out.println("ǰ�� ���� �����ڿ����ļ�:" + f1 + "  ��Ŀ���ļ��У�" + qd);
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

					if (placepath.indexOf("���") != -1) {
						System.out.println("��� ���� �����ڿ����ļ�:" + f1 + "  ��Ŀ���ļ��У�" + f2);
						fileChannelCope(f1, f2);
						isCopyed = true;
					}

				}

				if (!isCopyed) {
					fileChannelCope(f1, f2);
					System.out.println("���� �����ڿ����ļ�:" + f1 + "  ��Ŀ���ļ��У�" + f2);
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				System.err.println("�����ļ�����");
				e.printStackTrace();
			}
		}
	}

	/**
	 * excel �ļ�·��
	 */
	public String getExcelFilePath() {
		return excelFilePath;
	}

	/**
	 * Class �ļ�·��ǰ׺ �� com/javahis
	 */
	public String getClassPrefixPath() {
		return classPrefixPath;
	}

	/**
	 * ����ļ�û����չ��ʱ�Զ�׷�Ӹ���չ��
	 */
	public String getExpandedName() {
		return expandedName;
	}

	/**
	 * excel �ļ�·��
	 */
	public void setExcelFilePath(String excelFilePath) {
		this.excelFilePath = excelFilePath;
	}

	/**
	 * excel �ļ��и��°汾��(ͳһά��)�ֶ�
	 */
	public String getVersionNumber() {
		return versionNumber;
	}

	/**
	 * excel �ļ��и��°汾��(ͳһά��)�ֶ�
	 */
	public void setVersionNumber(String versionNumber) {
		this.versionNumber = versionNumber;
	}

	/**
	 * Class �ļ�·��ǰ׺ �� com/javahis
	 */
	public void setClassPrefixPath(String classPrefixPath) {
		this.classPrefixPath = classPrefixPath;
	}

	/**
	 * ����ļ�û����չ��ʱ�Զ�׷�Ӹ���չ��
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
	 * �ļ�����
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
		// File f = new File("F:/��װ/sjsk2.avi");
		// File f1 = new File("c:/sjsk2.avi");
		// try {
		// CopyFromExcel.fileChannelCope(f, f1);
		// } catch (Exception e) {
		// // TODO Auto-generated catch block
		// e.printStackTrace();
		// }
	}

	/**
	 * ����Excel���Ĺ�����
	 * 
	 * @author��xueyf
	 * @version 1.0
	 */
	class ExcelReader {
		private POIFSFileSystem fs;
		private HSSFWorkbook wb;
		private HSSFSheet sheet;
		private HSSFRow row;

		/**
		 * ��ȡExcel����ͷ������
		 * 
		 * @param InputStream
		 * @return String ��ͷ���ݵ�����
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
			// ����������
			int colNum = row.getPhysicalNumberOfCells();
			String[] title = new String[colNum];
			for (int i = 0; i < colNum; i++) {
				title[i] = getStringCellValue(row.getCell((short) i));
			}
			return title;
		}

		/**
		 * ��ȡExcel��������
		 * 
		 * @param InputStream
		 * @return Map ������Ԫ���������ݵ�Map����
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
			// �õ�������
			int rowNum = sheet.getLastRowNum();
			row = sheet.getRow(0);
			// ��������Ӧ�ôӵڶ��п�ʼ,��һ��Ϊ��ͷ�ı���
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

					// ׷���ļ�ǰ׺·��
					str = getClassPrefixPath() == null || getClassPrefixPath().equals("") ? str
							: getClassPrefixPath() + str;
					// б�ܵ���б�ܵ��滻
					str = str.replaceAll("\\\\", "\\/");
					// ׷���ļ���չ��
					str = str.lastIndexOf(".") == -1 ? str + (getExpandedName() == null ? "" : getExpandedName()) : str;
					// �ļ���չ���Ƚϣ����뵱ǰ������չ����ͬ�����追��
					if ((str.split("\\.").length == 1)
							|| (getCopyType() != null && !str.substring(str.lastIndexOf(".")).equals(getCopyType()))) {
						continue;
					}
				} catch (Exception ex) {
					System.out.println("����ѯ��" + i + "�������ļ�");
					ex.printStackTrace();
					break;
				}
				j++;
				count++;
				str = str + "@" + getStringCellValue(this.row.getCell(2)) + " ";
				content.put(i, str);

			}
			System.out.println("����ѯ��" + count + "�������ļ�");
			return content;
		}

		/**
		 * ��ȡExcel��������
		 * 
		 * @param InputStream
		 * @return Map ������Ԫ���������ݵ�Map����
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
			// �õ�������
			int rowNum = sheet.getLastRowNum();
			row = sheet.getRow(0);
			int colNum = row.getPhysicalNumberOfCells();
			// ��������Ӧ�ôӵڶ��п�ʼ,��һ��Ϊ��ͷ�ı���
			for (int i = 1; i <= rowNum; i++) {
				row = sheet.getRow(i);
				int j = 0;
				while (j < colNum) {
					// ÿ����Ԫ�������������"-"�ָ���Ժ���Ҫʱ��String���replace()������ԭ����
					// Ҳ���Խ�ÿ����Ԫ����������õ�һ��javabean�������У���ʱ��Ҫ�½�һ��javabean
					str += getStringCellValue(row.getCell((short) j)).trim() + "-";
					j++;
				}
				content.put(i, str);
				str = "";
			}
			return content;
		}

		/**
		 * ��ȡ��Ԫ����������Ϊ�ַ������͵�����
		 * 
		 * @param cell Excel��Ԫ��
		 * @return String ��Ԫ����������
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
		 * ��ȡ��Ԫ����������Ϊ�������͵�����
		 * 
		 * @param cell Excel��Ԫ��
		 * @return String ��Ԫ����������
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
					result = date.replaceAll("[����]", "-").replace("��", "").trim();
				} else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
					result = "";
				}
			} catch (Exception e) {
				System.out.println("���ڸ�ʽ����ȷ!");
				e.printStackTrace();
			}
			return result;
		}

	}

}
