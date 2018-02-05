### 关于GRID++的Java C\S调用实例
> 基于jacob-1.14.3的实现  

##### 1. 运行环境准备
1. 首先从GRID++官网[http://www.rubylong.cn/](http://www.rubylong.cn/)下载最新的GRID++编辑器,以及插件 
2. 拷贝gregn6.dll到jvm的bin目录下
3. 拷贝jacob-1.14.3.dll到jvm的bin目录下  

##### 2. 方法说明
| 方法名称        | 方法说明    | 
| --------       | -----   |
| regsvr32DLL    | 注册GRID++COM组件,在new此方法对象的时候会默认调用,采用CMD方式来进行注册|  
| execCommand   | 运行系统CMD命令|  
| setPrinterName| 设置打印机名称,如果不调用此方法名称,采用默认的打印机名称,如果调用此方法设置打印机名称,那么采用设置的打印机名称来使用   |
|getPrinterName | 获取当前的打印机名称,可以获取当前设置的打印机名称,以及设置后的打印机名称|
|AbortExport    |中断导出任务的执行,如果要导出PDF或则HTML,EXCEL等等,如果时间较长可以采用此方法进行中断导出的任务执行|
|AbortPrint     |调用此方法可以中断打印任务|
|About          |弹出GRID++的关于对话框,此对话框里面包含了GRID++的信息|
|AddParameter   |添加GRID++的参数|
|BeginLoopPrint |准备开始循环打印。|
|EndLoopPrint|结束循环打印|
|CancelLoadData |取消向报表载入数据所进行的准备任务。|
|Clear          |清除所有的报表模板定义数据，使报表主对象成为一张空白报表。|
|DeleteDetailGrid |  删除整个明细网格的定义，报表主对象将不具有明细网格。|
|DeletePageFooter | 删除页脚的定义，报表主对象将不具有页脚。|
|DeletePageHeader|删除页眉的定义，报表主对象将不具有页眉。|
|ExportDirect| 直接进行数据导出|
|ExtractXMLFromURL|从指定的 URL 获取XML数据包。|
|ExtractXMLFromURLEx|从指定的 URL 获取XML数据包。|
|ForceNewPage|在报表生成过程中，可以调用此方法强制进行分页。只能在PageProcessRecord 与 SectionFormat 事件中调用，在 SectionFormat 中调用时必须保证是在打印生成状态中。|
|GenerateDocumentFile|生成报表的 Grid++Report 文档文件到指定的文件。|
|LoadBackImageFromFile|从指定的图像文件中载入报表设计页面背景图。|
|LoadBackImageFromMemory|从内存数据中载入报表设计页面背景图。|
|LoadDataFromAjaxRequest|从内存数据中载入报表设计页面背景图。|
|LoadDataFromURL|从指定的 URL 地址载入报表明细记录集数据，数据必须为 XML 或 JSON 格式并符合约定的形式。|
|LoadDataFromURLEx|指定的 URL 地址载入报表数据，返回数据必须为 XML 或 JSON 格式并符合约定的形式。|
|LoadDataFromXML|从 XML 或 JSON 文字串中载入报表明细记录集数据，数据应符合约定的形式。|
|LoadFromFile|从报表模板文件中载入报表模板定义。|
|LoadFromMemory|从内存中载入报表模板数据。   |
|LoadFromStr|从字符串中载入报表模板数据|
|LoadFromURL|从指定的 URL 地址载入报表模板数据。|
|LoadWatermarkFromFile|从指定的图像文件中载入水印背景图。|
|LoadWatermarkFromMemory| 从内存数据中载入报表水印背景图。  |
|LoopPrint|在循环打印过程中，执行一次打印生成过程。|
|PixelsToUnit|将屏幕像素值转换为报表当前使用的计量单位值。|
|PrepareLoadData|准备向报表中载入数据，调用此方法后，就可以向报表的记录集推送(加入)数据。|
|Print|执行打印报表。   |
|PrintEx|打印报表,按指定方式生成生成打印数据|
|PrintPreview|在Grid++Report提供的缺省打印预览窗口中预览报表。  |
|PrintPreviewEx| 在Grid++Report提供的缺省打印预览窗口中预览报表，可以指定打印数据的生成方式。|
|Register|用Grid++Report的产品序列号进行注册。|
|SaveToFile|将报表模板数据保存到文件中，以模板设置的存储格式保存模板数据。|
|Release|释放资源|
|fileToBytes|通过文件路径,将文件转换为字节|  

```java
package com.grid.models;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Before;
import org.junit.Test;


public class TestGridUtils {
	private GridReportUtils gru;
	private String filepath = "C:\\\\zero_bulk.grf";
	private List<Object> datas =new ArrayList<Object>();
	@Before 
	public void bootstrap() {
		gru = new GridReportUtils();
		/***创建测试数据***/
		Map<String,String> datarow = new HashMap<String,String>();
		datarow.put("PACKINGQTY", "12");
		datas.add(datarow);
	}
	/**
	 * 测试直接打印
	 */
	@Test
	public void testDirectPrint() {
		gru.LoadFromFile(filepath);
		gru.LoadDataJavaListDatas(datas);
		gru.Print(false); //true表示不显示对话框,false表示显示对话框
	}
	/**
	 * 测试是否弹出grid的对话框
	 */
	@Test
	public void testAbout() {
		gru.About();
	}
	/**
	 * 预览打印
	 * */
	@Test
	public void testPrintPreview() {
		gru.LoadFromFile(filepath);
		gru.LoadDataJavaListDatas(datas);
		gru.PrintPreview(true); //true表示不显示对话框,false表示显示对话框
	}
}

``` 


使用  gru.LoadFromFile(filepath);来装载模板文件  
使用  gru.LoadDataJavaListDatas(datas); 来装载数据
使用  gru.PrintPreview(true);或gru.Print(false)来进行打印或者预览打印

部分方法为来的及进行测试,只是测试部分代码如果有BUG或者疑问,请在github上给我提issue

