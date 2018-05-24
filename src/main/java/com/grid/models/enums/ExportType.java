package com.grid.models.enums;
/**
 * 指定数据导出的类型。
 * @author zhangj
 * @date   2018-05-23 22:12:34
 * @email  zhangjin0908@hotmail.com
 */
public enum ExportType {
	/**
	 * 导出Excel文件。
	 */
	gretXLS(1),
	/**
	 * 导出文本文件。
	 */
	gretTXT(2),
	/**
	 * 导出Html超文本文件。
	 */
	gretHTM(3),
	/**
	 * 导出RTF文件。
	 */
	gretRTF(4),
	/**
	 * 导出PDF格式文件。
	 */
	gretPDF(5),
	/**
	 * 导出CSV格式文件。
	 */
	gretCSV(6),
	/**
	 * 导出图像文件，支持多种图像格式。
	 */
	gretIMG(7);
	private int value;
	
	public int getValue() {
		return value;
	}

	ExportType(int value) {
		this.value=value;
	}
}
