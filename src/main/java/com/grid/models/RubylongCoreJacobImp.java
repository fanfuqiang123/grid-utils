package com.grid.models;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.grid.models.enums.ExportType;
import com.grid.models.enums.GenerateStyle;
import com.grid.models.enums.GridVersion;
import com.grid.models.enums.ImageType;
import com.grid.models.interfaces.Json;
import com.grid.models.interfaces.RubylongCore;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 
 * @author zhangj
 * @date 2018-05-23 21:49:36
 * @email zhangjin0908@hotmail.com
 */
public class RubylongCoreJacobImp  implements RubylongCore{
	private static String VERSION = "6.5";
	private static String PROGRAMID = "gregn.GridppReport.6";
	
	private ActiveXComponent com;
	private Dispatch disp;
	public static String getVERSION() {
		return VERSION;
	}

	public static String getPROGRAMID() {
		return PROGRAMID;
	}

	public static void setPROGRAMID(String pROGRAMID) {
		PROGRAMID = pROGRAMID;
	}
	

	protected RubylongCoreJacobImp(GridVersion version) {
		ComThread.InitSTA();
		com = new ActiveXComponent("gregn.GridppReport.6");
		disp = com.getObject();
	}

	@Override
	public String getVersion() {
		return VERSION;
	}
	@Override
	public void setPrinterName(String printerName) {
		Dispatch d = Dispatch.get(disp, "Printer").getDispatch();
		Dispatch.put(d, "PrinterName", printerName);
	}
	@Override
	public String getPrinterName() {
		Dispatch d = Dispatch.get(disp, "Printer").getDispatch();
		return Dispatch.get(d, "PrinterName").getString();
	}
	@Override
	public void release() {
		ComThread.Release();
	}

	@Override
	public void fnPrintPreview(boolean showModal) {
		Dispatch.call(disp, "PrintPreview",new Variant(showModal));
	}

	@Override
	public void fnPrintEx(GenerateStyle generateStyle, boolean showPrintDialog) {
		 Dispatch.call(disp, "PrintEx",new Variant(generateStyle.getValue()),new Variant(showPrintDialog));
	}

	@Override
	public void fnPrint(boolean showPrintDialog) {
		 Dispatch.call(disp, "Print",new Variant(showPrintDialog));
	}

	
	@Override
	public boolean fnLoadDataFromURL(String url) {
		Variant v = Dispatch.call(disp, "LoadDataFromURL",new Variant(url));
		return v.getBoolean();
	}
	
	@Override
	public boolean fnLoadDataFromXML(String data) {
		Variant v = Dispatch.call(disp, "LoadDataFromXML",new Variant(data));
		return v.getBoolean();
	}

	@Override
	public void fnLoadDataListDatas(List<?> datas,Json json) {
		Map<String,Object> map=new HashMap<String,Object>();
		map.put("Detail", datas);
		this.fnLoadDataFromXML(json.toJSONString(map));
	}

	@Override
	public void fnExportImg(ImageType imageType, String fileName, boolean allInOne, boolean doneOpen){
		Dispatch ExportOption = Dispatch.call(disp, "PrepareExport", new Variant(ExportType.gretIMG.getValue())).getDispatch();
		Dispatch E2IMGOption = Dispatch.get(ExportOption, "AsE2IMGOption").getDispatch();
		Dispatch.put(E2IMGOption, "ImageType", new Variant(imageType.getValue()));
		Dispatch.put(E2IMGOption, "AllInOne", new Variant(allInOne));
		Dispatch.put(E2IMGOption, "FileName", new Variant(fileName));
		Dispatch.call(disp, "Export", new Variant(doneOpen));
		Dispatch.call(disp, "UnprepareExport");
	}


	@Override
	public boolean fnExportDirect(ExportType exportType, String fileName, boolean showOptionDlg, boolean doneOpen) {
		Variant v = Dispatch.call(disp, "ExportDirect",new Variant(exportType.getValue()),new Variant(fileName),new Variant(showOptionDlg),new Variant(doneOpen));
		return v.getBoolean();
	}

	@Override
	public boolean fnLoadFromURL(String url) {
		Variant v = Dispatch.call(disp, "LoadFromURL",new Variant(url));
		return v.getBoolean();
	}
	
	@Override
	public boolean fnLoadFromFile(String fileName) {
		Variant v = Dispatch.call(disp, "LoadFromFile",new Variant(fileName));
		return v.getBoolean();		
	}
	
	@Override
	public boolean fnLoadFromStr(String template) {
		Variant v = Dispatch.call(disp, "LoadFromStr",new Variant(template));
		return v.getBoolean();		
	}

	@Override
	public void about() {
		Dispatch.call(disp, "About");		
	}

	private static RubylongCore rlc ;

	public static RubylongCore getRubylongCore(GridVersion gridVersion) {
		if(rlc==null) {
			rlc=new RubylongCoreJacobImp(gridVersion);
		}
		return rlc;
	}
	

}
