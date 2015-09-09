package com.mani.app;

import generated.Shiporder;

public class Main {

	public static void main(String[] args) {
		Shiporder data = ExcelParser.parse("D:\\manikanta\\Experiments\\Sample-excel-Input.xlsx");
		String xmlData = CustomParser.marshall(data);
		System.out.println(xmlData);
	}
}
