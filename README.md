### 关于GRID++的Java C\S调用实例
> 基于jacob-1.14.3的实现  

##### 1. 运行环境准备 
__依赖__:net.sf.jacob-project:jacob:1.14.3 

1. 首先从GRID++官网`http://www.rubylong.cn/`下载最新的GRID++编辑器,以及插件   
2. 拷贝jacob-1.14.3.dll `https://sourceforge.net/projects/jacob-project/files/jacob-project/1.14.3/` 到jvm的bin目录下,让dll加载成功
> 环境准备好了,运行TestRubylongCoreJacobImp测试类进行测试


##### 2. 方法说明
| 方法名称        | 方法说明    |    
| --------       | -----   |    
|setPrinterName| 设置打印机名称,如果不调用此方法名称,采用默认的打印机名称,如果调用此方法设置打印机名称,那么采用设置的打印机名称来使用   |  
|getPrinterName | 获取当前的打印机名称,可以获取当前设置的打印机名称,以及设置后的打印机名称|   
|About          |弹出GRID++的关于对话框,此对话框里面包含了GRID++的信息|   
|LoadDataFromXML|从 XML 或 JSON 文字串中载入报表明细记录集数据，数据应符合约定的形式。|  
|LoadFromFile|从报表模板文件中载入报表模板定义。|  
|LoadFromMemory|从内存中载入报表模板数据。   |  
|LoadFromStr|从字符串中载入报表模板数据|  
|Print|执行打印报表。   |  
|PrintEx|打印报表,按指定方式生成生成打印数据|  
|PrintPreview|在Grid++Report提供的缺省打印预览窗口中预览报表。  |  
|Release|释放资源|  
|getVersion|获取当前对应的Grid++Report的版本


> 最近大家有问题都是发的邮件，是否可以在github里面编辑iess这样大家遇到的问题，其他人也可以进行一起讨论

#### 如果有问题，可以联系Wechat或者手机号码 - 

Wechat/手机号码 Base64： MTg4MjcwMDU2OTQ=  
