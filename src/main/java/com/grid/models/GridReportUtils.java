package com.grid.models;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.alibaba.fastjson.JSON;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class GridReportUtils {
	private Dispatch disp;
	private String defaultPrinter;
	public GridReportUtils() {
		ComThread.InitSTA();
		ActiveXComponent com  =null;
		try {
			com = new ActiveXComponent("grpro.GridppReport.6");
		} catch (com.jacob.com.ComFailException e) {
			try {
				regsvr32DLL();
				Thread.sleep(5000);
				com = new ActiveXComponent("grpro.GridppReport.6");
			} catch (IOException e1) {
				e1.printStackTrace();
			} catch (InterruptedException e1) {
				e1.printStackTrace();
			}
		}
		disp = com.getObject();
		defaultPrinter=getPrinterName();
	}
	static public void regsvr32DLL() throws IOException {
		execCommand("regsvr32  "+System.getProperty("java.home")+"\\bin\\gregn6.dll");
	}
	static public void execCommand(String commandStr) throws IOException {
		Runtime.getRuntime().exec(commandStr);
	}
	/**
	 * 设置打印机名称
	 * @param PrinterName 要设置的默认打印机的名称
	 */
	public void setPrinterName(String PrinterName) {
		 Dispatch d = Dispatch.get(disp, "Printer").getDispatch();
		 Dispatch.put(d, "PrinterName", PrinterName);
	}

	/**
	 * 获取当前打印机的名称
	 * @return
	 */
	public String getPrinterName() {
		Dispatch d = Dispatch.get(disp, "Printer").getDispatch();
		return Dispatch.get(d, "PrinterName").getString();
	}
	/**
	 * 中断数据导出任务的执行
	 * 此方法只能在 ExportBegin 事件中执行，用于中止数据导出
	 */
	public void AbortExport() {
		Dispatch.call(disp, "AbortExport");
	}
	/**
	 * 中断打印任务的执行
	 * 此方法只能在 PrintBegin 与 PrintPage 事件中执行。
	 */
	public void AbortPrint() {
		Dispatch.call(disp, "AbortPrint");
	}
	
	/**
	 * 显示Grid++Report关于对话框
	 */
	public void About() {
		Dispatch.call(disp, "About");
	}
	/**
	 * 增加一个参数   
	 * @param name  指定参数的名称标识
	 * @param type  指定参数的类型
	 *  grptString 1 字符串类型，可以为任意长度。 
	 *	grptInteger 2 整数类型，可以设定格式化串。 
	 *	grptFloat 3 浮点数类型，可以设定格式化串。 
	 *	grptBoolean 5 布尔类型，可以真值与假值的显示文字。 
	 *	grptDateTime 6 日期时间类型，可以设定格式化串。 
	 * @return
	 */
	public Variant AddParameter(String name ,int type) {
		return Dispatch.call(disp, "AddParameter",new Variant(name),new Variant(type));
	}
	// 
	/**
	 * 准备开始循环打印。
	 * @param generateStyle  
	 * 	grpgsOnlyForm 1 只生成报表表单数据。 
	 *	grpgsOnlyContent 2 只生成报表内容数据。 
	 *	grpgsAll 3 生成报表所有数据。 
	 *	grpgsPreviewAll 4 预览报表全部内容，但只打印出内容数据。 
     *
	 * @param showPrintDlg
	 */
	public void BeginLoopPrint(int generateStyle,boolean showPrintDlg) {
		 Dispatch.call(disp, "BeginLoopPrint",new Variant(generateStyle),new Variant(showPrintDlg));
	}
	
	//CalcTextOut  技术有限暂时不知道如何翻译此方法
	
	/**
	 * 取消向报表载入数据所进行的准备任务。
	 */
	public void CancelLoadData() {
		 Dispatch.call(disp, "CancelLoadData");
	}
	
	/**
	 * 清除所有的报表模板定义数据，使报表主对象成为一张空白报表。
	 */
	public void Clear() {
		Dispatch.call(disp, "Clear");
	}
	
	//ColumnByName  技术有限暂时不知道如何实现此方法
	
	//ControlByName 技术有限暂时不知道如何实现此方法
	
	/**
	 * 删除整个明细网格的定义，报表主对象将不具有明细网格。
	 */
	public void DeleteDetailGrid() {
		Dispatch.call(disp, "DeleteDetailGrid");
	}
	/**
	 * 删除页脚的定义，报表主对象将不具有页脚。
	 */
	public void DeletePageFooter() {
		Dispatch.call(disp, "DeletePageFooter");
	}
	/**
	 * 删除页眉的定义，报表主对象将不具有页眉。
	 */
	public void DeletePageHeader() {
		Dispatch.call(disp, "DeletePageHeader");
	}
	/**
	 * 结束循环打印
	 */
	public void EndLoopPrint() {
		Dispatch.call(disp, "EndLoopPrint");
	}
	/**
	 * 直接进行数据导出。
	 * @param ExportType  指定导出的文件类型。
	 *  gretXLS 1 导出Excel文件。 
	 *  gretTXT 2 导出文本文件。 
	 *  gretHTM 3 导出Html超文本文件。 
	 *	gretRTF 4 导出RTF文件。 
	 *	gretPDF 5 导出PDF格式文件。 
	 *	gretCSV 6 导出CSV格式文件。 
	 *	gretIMG 7 导出图像文件，支持多种图像格式。 
	 * @param FileName       指定导出的完整文件路径与文件名称。
	 * @param ShowOptionDlg  指定是否在导出之前显示选项设置对话框。 
	 * @param DoneOpen       指示是否在导出数据之后用关联程序打开导出文件。
	 * @return  指示是否成功进行了数据导出。
	 */
	public boolean ExportDirect(int ExportType, String FileName, boolean ShowOptionDlg, boolean DoneOpen) {
		Variant v = Dispatch.call(disp, "ExportDirect",new Variant(ExportType),new Variant(FileName),new Variant(ShowOptionDlg),new Variant(DoneOpen));
		return v.getBoolean();
	}
	/**
	 * 从指定的 URL 获取XML数据包。 
	 * @param url  获得的 XML 数据文本。
	 * @return 获取数据的 URL。 
	 */
	public String ExtractXMLFromURL(String url) {
		Variant v = Dispatch.call(disp, "ExtractXMLFromURL",new Variant(url));
		return v.getString();
	}
	/**
	 *  从指定的 URL 获取XML数据包。
	 * @param DataURL    获取数据的 URL。 
	 * @param DataParam  向服务器提供的参数。 
	 * @return  获得的 XML 数据文本。
	 */
	public String ExtractXMLFromURLEx(String DataURL, String DataParam) {
		Variant v = Dispatch.call(disp, "ExtractXMLFromURLEx",new Variant(DataURL),new Variant(DataParam));
		return v.getString();
	}
	
	
	//FieldByDBName    技术有限暂时不知道如何实现此方法
	//FieldByName      技术有限暂时不知道如何实现此方法
	//FindFirstControl 技术有限暂时不知道如何实现此方法
	//FindNextControl  技术有限暂时不知道如何实现此方法
	
	/**
	 * 在报表生成过程中，结束当前页，开始一个新页面。
	 * 在报表生成过程中，可以调用此方法强制进行分页。只能在PageProcessRecord 与 SectionFormat 事件中调用，在 SectionFormat 中调用时必须保证是在打印生成状态中。
	 */
	public void ForceNewPage() {
		Dispatch.call(disp, "ForceNewPage");
	}
	
	//GenerateDocumentData  技术有限暂时不知道如何实现此方法
	
	/**
	 * 生成报表的 Grid++Report 文档文件到指定的文件。
	 * @param PathFile  保存数据的文件的路径与名称。
	 */
	public void GenerateDocumentFile(String PathFile) {
		Dispatch.call(disp, "GenerateDocumentFile",new Variant(PathFile));
	}
	//InsertDetailGrid      技术有限暂时不知道如何实现此方法
	//InsertPageFooter      技术有限暂时不知道如何实现此方法
	//InsertPageHeader      技术有限暂时不知道如何实现此方法
	
	/**
	 * 从指定的图像文件中载入报表设计页面背景图。
	 * @param PathFile 指定图像文件的路径与名称。 
	 */
	public void LoadBackImageFromFile(String PathFile) {
		Dispatch.call(disp, "LoadBackImageFromFile",new Variant(PathFile));
	}
	
	/**
	 * 从内存数据中载入报表设计页面背景图。
	 * @param pBuffer    指定图像数据的内存地址。 
	 * @param BytesCount 指定图像数据的内存字节数。 
	 */
	public void LoadBackImageFromMemory(byte[] pBuffer, int BytesCount) {
		Dispatch.call(disp, "LoadBackImageFromMemory",new Variant(pBuffer),new Variant(BytesCount));
	}
	/**
	 * 将WEB报表展现网页中用 Ajax XMLHttpRequest 对象请求到的数据加载到报表中。
	 * @param ResponseText    HTTP响应数据体文本。 
	 * @param ResponseHeaders HTTP响应数据头文本。 
	 */
	public void LoadDataFromAjaxRequest(String ResponseText, String ResponseHeaders) {
		Dispatch.call(disp, "LoadDataFromAjaxRequest",new Variant(ResponseText),new Variant(ResponseHeaders));
	}
	
	/**
	 * 从指定的 URL 地址载入报表明细记录集数据，数据必须为 XML 或 JSON 格式并符合约定的形式。
	 * @param URL  载入数据的URL。 
	 * @return  指定载入是否成功
	 */
	public boolean LoadDataFromURL(String URL) {
		Variant v = Dispatch.call(disp, "LoadDataFromURL",new Variant(URL));
		return v.getBoolean();
	} 
	
	/**
	 * 指定的 URL 地址载入报表数据，返回数据必须为 XML 或 JSON 格式并符合约定的形式。
	 * @param DataURL     载入数据的URL。 
	 * @param DataParam   向服务器提供的额外参数
	 * @return   指定载入是否成功。
	 */
	public boolean LoadDataFromURLEx(String DataURL,String DataParam) {
		Variant v = Dispatch.call(disp, "LoadDataFromURL",new Variant(DataURL),new Variant(DataParam));
		return v.getBoolean();
	} 
	
	/**
	 * 从 XML 或 JSON 文字串中载入报表明细记录集数据，数据应符合约定的形式。
	 * @param XMLData   XML数据文本。
	 * @return   指示载入是否成功。
	 */ 
	public boolean LoadDataFromXML(String XMLData) {
		Variant v = Dispatch.call(disp, "LoadDataFromXML",new Variant(XMLData));
		return v.getBoolean();
	} 
	/**
	 * 从报表模板文件中载入报表模板定义。
	 * @param FileName   指定模板文件的完全路径名称
	 * @return   指示载入报表模板数据是否成功。
	 */
	public boolean LoadFromFile(String FileName) {
		Variant v = Dispatch.call(disp, "LoadFromFile",new Variant(FileName));
		return v.getBoolean();
	}
	/**
	 * 从内存中载入报表模板数据。   
	 * @param pData      数据地址指针 
	 * @param ByteCount  数据字节大小 
	 * @return   指示载入报表模板数据是否成功
	 */
	public boolean LoadFromMemory(byte[] pData,int ByteCount) {
		Variant v = Dispatch.call(disp, "LoadFromMemory",new Variant(pData),new Variant(ByteCount));
		return v.getBoolean();
	}
	/**
	 * 从字符串中载入报表模板数据
	 * @param Data   指定报表模板数据。 
	 * @return 指示载入报表模板数据是否成功。
	 */
	public boolean LoadFromStr(String Data) {
		Variant v = Dispatch.call(disp, "LoadFromStr",new Variant(Data));
		return v.getBoolean();
	}
	/**
	 * 从指定的 URL 地址载入报表模板数据。
	 * @param URL   载入的URL地址。 
	 * @return  指示载入是否成功。 
	 */
	public boolean LoadFromURL(String URL) {
		Variant v = Dispatch.call(disp, "LoadFromURL",new Variant(URL));
		return v.getBoolean();
	}
	
	//LoadFromVariant 技术有限暂时不知道如何实现此方法
	/**
	 * 从指定的图像文件中载入水印背景图。
	 * @param PathFile  指定图像文件的路径与文件名。
	 */
	public void LoadWatermarkFromFile(String PathFile) {
		Dispatch.call(disp, "LoadWatermarkFromFile",new Variant(PathFile));
	}
	/**
	 * 从内存数据中载入报表水印背景图。  
	 * @param pBuffer     指定图像数据的内存地址。 
	 * @param BytesCount  指定图像数据的内存字节数
	 */
	public void LoadWatermarkFromMemory(byte[] pBuffer,int BytesCount ) {
		Dispatch.call(disp, "LoadWatermarkFromFile",new Variant(pBuffer),new Variant(BytesCount));
	}
	
	/***
	 * 在循环打印过程中，执行一次打印生成过程。
	 */
	public void LoopPrint() {
		Dispatch.call(disp, "LoopPrint");
	}
	
	//ParameterByName  技术有限暂时不知道如何实现此方法
	
	/**
	 * 将屏幕像素值转换为报表当前使用的计量单位值。
	 * @param Pixels   指定要转换的屏幕像素值。 
	 * @return   转换得到的计量单位值。
	 */
	public double PixelsToUnit(int Pixels) {
		Variant v = Dispatch.call(disp, "PixelsToUnit",new Variant(Pixels));
		return v.getDouble();
	}
	
	//PrepareExport   技术有限暂时不知道如何实现此方法
	
	/**
	 * 准备向报表中载入数据，调用此方法后，就可以向报表的记录集推送(加入)数据。
	 * @return  指示调用是否成功
	 */
	public boolean PrepareLoadData() {
		Variant v = Dispatch.call(disp, "PrepareLoadData");
		return v.getBoolean();
	}
	
	/**
	 * 执行打印报表。   
	 * @param ShowPrintDialog  指示在打印前是否显示打印对话框。 
	 */
	public void Print(boolean ShowPrintDialog) {
		 Dispatch.call(disp, "Print",new Variant(ShowPrintDialog));
	}
	
	//PrintDocumentData  技术有限暂时不知道如何实现此方法
	
	/**
	 * 打印报表,按指定方式生成生成打印数据
	 * @param GenerateStyle   指定打印数据的生成类型。 
	 *  grpgsOnlyForm 1 只生成报表表单数据。 
	 *	grpgsOnlyContent 2 只生成报表内容数据。 
	 *	grpgsAll 3 生成报表所有数据。 
	 *	grpgsPreviewAll 4 预览报表全部内容，但只打印出内容数据。 
	 * @param ShowPrintDialog  指示在打印前是否显示打印对话框。 
	 */
	public void PrintEx(int GenerateStyle,boolean ShowPrintDialog) {
		 Dispatch.call(disp, "PrintEx",new Variant(GenerateStyle),new Variant(ShowPrintDialog));
	}
	
	/**
	 * 在Grid++Report提供的缺省打印预览窗口中预览报表。  
	 * @param ShowModal    指定是否以有模的方式显示窗口
	 */
	public void PrintPreview(boolean ShowModal) {
		Dispatch.call(disp, "PrintPreview",new Variant(ShowModal));
	}
	
	/**
	 * 在Grid++Report提供的缺省打印预览窗口中预览报表，可以指定打印数据的生成方式。
	 * @param GenerateStyle    指定打印数据的生成方式。 
	 *	grpgsOnlyForm 1 只生成报表表单数据。 
	 *	grpgsOnlyContent 2 只生成报表内容数据。 
	 *	grpgsAll 3 生成报表所有数据。 
	 *	grpgsPreviewAll 4 预览报表全部内容，但只打印出内容数据。 
	 * @param ShowModal        指定是否以有模的方式显示打印预览窗口。 
	 * @param MaximizeWindow   指定是否最大化显示打印预览窗口。 
	 */
	public void PrintPreviewEx(int GenerateStyle, boolean ShowModal, boolean MaximizeWindow) {
		Dispatch.call(disp, "PrintPreviewEx",new Variant(GenerateStyle),new Variant(ShowModal),new Variant(MaximizeWindow));
	}
	/**
	 * 用Grid++Report的产品序列号进行注册。
	 * @param SerialNo  产品序列号 
	 * @return  指示注册是否成功
	 */
	public boolean  Register(String SerialNo) {
		Variant v = Dispatch.call(disp, "Register",new Variant(SerialNo));
		return v.getBoolean();
	}
	/**
	 * 将报表模板数据保存到文件中，以模板设置的存储格式保存模板数据。
	 * @param FileName  指定模板文件的完全路径名称。
	 * @return  指示保存报表模板数据是否成功。
	 */
	public boolean SaveToFile(String FileName) {
		Variant v = Dispatch.call(disp, "SaveToFile",new Variant(FileName));
		return v.getBoolean();
	}
	
	/**
	 * 释放资源
	 */
	public void Release() {
		ComThread.Release();
	}
	/**
	 * 装载数据到Grid++表格
	 * @param datas 
	 */
	public void LoadDataJavaListDatas(List<Object> datas) {
		Map<String,Object> map=new HashMap<String,Object>();
		map.put("Detail", datas);
		this.LoadDataFromXML(JSON.toJSONString(map));
	}
	public static byte[] fileToBytes(String filepath) throws IOException {
		File file = new File(filepath);
		InputStream input =  null;
		try {
			input=new FileInputStream(file);
			byte[] byt = new byte[input.available()];
			input.read(byt);
			return byt;
		} finally {
			if(input!=null) {
				input.close();
			}
		}
			
	}	
}
