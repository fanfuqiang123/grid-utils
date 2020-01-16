package com.grid.models.enums;

/**
 * 指定在导出图像文件时可用的各种图像格式
 */
public enum ImageType {

	greitBMP(1),

	greitPNG(2),

	greitJPEG(3),

	greitTIFF(5);

	private int value;

	public int getValue() {
		return value;
	}

	ImageType(int value) {
		this.value = value;
	}
}
