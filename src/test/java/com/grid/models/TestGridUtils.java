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
