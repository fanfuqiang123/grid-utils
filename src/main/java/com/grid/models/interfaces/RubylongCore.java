package com.grid.models.interfaces;

import java.util.List;

import com.grid.models.enums.ExportType;
import com.grid.models.enums.GenerateStyle;
import com.grid.models.enums.ImageType;

public interface RubylongCore {
	/**
	 * 获取当前程序对应的瑞浪报表的版本
	 * @return 返回对应的版本信息
	 */
	String getVersion();
	/**
	 * 释放COM资源,如果使用完毕后请释放COM资源,此资源不会被垃圾回收机制管理
	 */
	void   release();
	/**
	 * 在Grid++Report提供的缺省打印预览窗口中预览报表。
	 * @param showModal
	 *            指定是否以有模的方式显示窗口
	 */
	void  fnPrintPreview(boolean showModal);
	
	/**
	 * 打印报表,按指定方式生成生成打印数据
	 * 
	 * @param generateStyle
	 *            指定打印数据的生成类型。 grpgsOnlyForm 1 只生成报表表单数据。 grpgsOnlyContent 2
	 *            只生成报表内容数据。 grpgsAll 3 生成报表所有数据。 grpgsPreviewAll 4
	 *            预览报表全部内容，但只打印出内容数据。
	 * @param showPrintDialog
	 *            指示在打印前是否显示打印对话框。
	 */
	void  fnPrintEx(GenerateStyle generateStyle, boolean showPrintDialog);
	/**
	 * 执行打印报表。
	 * 
	 * @param showPrintDialog
	 *            指示在打印前是否显示打印对话框。
	 */
	void  fnPrint(boolean showPrintDialog);
	
	/**
	 * 从指定的 URL 地址载入报表明细记录集数据，数据必须为 XML 或 JSON 格式并符合约定的形式。
	 * @param url 载入数据的URL
	 * @return 指示载入报表数据是否成功。
	 */
	boolean  fnLoadDataFromURL(String url);

	/**
	 * 从字符串中载入报表数据
	 * @param data 指定报表数据。
	 * @return 指示载入报表数据是否成功。
	 */
	boolean  fnLoadDataFromXML(String data);
	
	/**
	 * 装载报表所需求的数据
	 * @param datas  要装载的数据
	 * @param json   要转换的json工具类
	 */
	void fnLoadDataListDatas(List<?> datas,Json json);
	
	/**
	 * 从指定的 URL 地址载入报表模板数据。
	 * @param url 载入的URL
	 * @return 指示载入是否成功。
	 */
	boolean  fnLoadFromURL(String url); // 有水印
	
	
	/**
	 * 从报表模板文件中载入报表模板定义
	 * @param fileName 指定模板文件的完全路径名称。 
	 * @return 指示载入是否成功。
	 */
	boolean fnLoadFromFile(String fileName);
	
	
	/**
	 * 装载报表模版
	 * @param template
	 * @return 指示载入报表模版是否成功。
	 */
	boolean fnLoadFromStr(String template);
	/**
	 * 直接进行数据导出。
	 * @param exportType  指定导出的文件类型。
	 * @param fileName       指定导出的完整文件路径与文件名称。
	 * @param showOptionDlg  指定是否在导出之前显示选项设置对话框。
	 * @param doneOpen       指示是否在导出数据之后用关联程序打开导出文件。
	 * @return  指示是否成功进行了数据导出。
	 */
	boolean fnExportDirect(ExportType exportType,String fileName, boolean showOptionDlg, boolean doneOpen);

	/**
	 * 导出图像文件
	 * @param imageType	指定在导出图像文件时可用的各种图像格式。
	 * @param fileName	指定导出的完整文件路径与文件名称。
	 * @param allInOne    指示是否将整个报表导出在一个图像文件中。
	 * @param doneOpen	指示是否在导出数据之后用关联程序打开导出文件。
	 */
	void fnExportImg(ImageType imageType, String fileName, boolean allInOne, boolean doneOpen);
	/**
	 * 设置打印机名称,如果设置这个打印机名称那么将会使用此打印机进行打印
	 * @param printName
	 */
	void setPrinterName(String printName);
	/**
	 * 获取当前设置的打印机名称
	 * @return
	 */
	String getPrinterName();
	/**
	 * 关于瑞浪报表的详细介绍
	 */
	void about();
	
}
