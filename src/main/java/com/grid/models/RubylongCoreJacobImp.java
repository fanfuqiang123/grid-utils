package com.grid.models;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.alibaba.fastjson.JSON;
import com.grid.models.enums.ExportType;
import com.grid.models.enums.GenerateStyle;
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
	

	public RubylongCoreJacobImp() {
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
	public boolean fnExportDirect(ExportType exportType, String fileName, boolean showOptionDlg, boolean doneOpen) {
		Variant v = Dispatch.call(disp, "ExportDirect",new Variant(exportType.getValue()),new Variant(fileName),new Variant(showOptionDlg),new Variant(doneOpen));
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

}
