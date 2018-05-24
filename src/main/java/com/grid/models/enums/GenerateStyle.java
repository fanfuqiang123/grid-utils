package com.grid.models.enums;
/**
 * 打印生成时，如果指定打印生成方式为只输出报表内容数据（OnlyContent）,则只输出打印类别为报表内容数据的报表元素，其它打印类别为报表表单数据的报表元素将不会输出。在只输出报表内容数据时，报表节的背景色不会填充，明细网格的边框线与行列线也不会输出。在只输出报表内容数据的打印生成方式特别适用于进行票据套打与现存格式报表的套打。
 * 打印生成时，如果指定印生成方式为输出报表表单数据（OnlyForm）,生成方式与只输出报表内容数据（OnlyContent）反之对应。在此种方式下，明细网格的记录数据不会被应用，明细行的生成与明细记录不对应，不填充明细记录页也可以输出报表表单数据。报表表单数据一般只生成一页。此种模式一般用来制作票据与固定格式报表的印刷模版。
 * 打印生成时，如果指定印生成方式（PrintGenerateStyle）为输出全部数据（All）,则所有数据都会输出，这也是我们平时最多用到的方式。
 * 使用打印页面的生成方式的地方有： 在IGridppReport接口中执行PrintEx与PrintViewEx方法中可以指定打印页面的生成方式，另可以设定打印预览查看器的打印页面的生成方式（GenerateStyle）属性
 * @author zhangj
 * @date   2018-05-23 22:02:17
 * @email  zhangjin0908@hotmail.com
 */
public enum GenerateStyle {
	/***
	 * 只生成报表表单数据。
	 */
	grpgsOnlyForm(1),
	/**
	 * 只生成报表内容数据。
	 */
	grpgsOnlyContent(2),
	/**
	 * 生成报表所有数据。
	 */
	grpgsAll(3),
	/**
	 * 预览报表全部内容，但只打印出内容数据。
	 */
	grpgsPreviewAll(4);
	private int value;
	GenerateStyle(int value){
		this.value=value;
	}
	public int getValue() {
		return value;
	}
	
}
