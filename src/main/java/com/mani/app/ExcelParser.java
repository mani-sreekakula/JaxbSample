package com.mani.app;

import generated.CommonData;
import generated.Grid;
import generated.Grids;
import generated.Orders;
import generated.Shiporder;
import generated.Template;
import generated.Templates;
import generated.Variable;
import generated.Variables;

import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelParser {

	public static void main(String[] args) {
		Shiporder data = parse("D:\\manikanta\\Experiments\\Sample-excel-Input.xlsx");
		System.out.println(data);
	}

	public static Shiporder parse(String fileName) {
		Shiporder shiporder = new Shiporder();
		try {
			InputStream inStream = new FileInputStream(fileName);
			List<Map<String, String>> data = parse(inStream);
			List<Template> templates = customizeToModel(data);
			CommonData commonData = getHardCodedData();
			Templates templates1  = new Templates();
			templates1.setTemplate(templates.get(0));
			Orders orders = new Orders();
			orders.setCommonData(commonData);
			orders.setTemplates(templates1);
			shiporder.setOrders(orders);
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		return shiporder;
	}

	private static CommonData getHardCodedData() {
		CommonData commonData = new CommonData();
		commonData.setTransactionID("123ABC");
		commonData.setData1("ACD");
		commonData.setData2("chef");
		commonData.setData3("aced");
		commonData.setData4("xys");
		commonData.setData5("ery");
		commonData.setData6("123");
		commonData.setData7("jks123");
		return commonData;
	}

	private static List<Template> customizeToModel(List<Map<String, String>> data) {
		List<Template> templates = new ArrayList<Template>();
		for (Map<String, String> record : data) {
			Template template = new Template();
			Grid grid = new Grid();
			Variable variable = new Variable();
			for (Iterator<String> it = record.keySet().iterator(); it.hasNext();) {
				String fieldName = it.next();
				String field = fieldName.substring(0, 1).toLowerCase() + fieldName.substring(1);
				setValueToField(template, field, record.get(fieldName));
				setValueToField(grid, field, record.get(fieldName));
				setValueToField(variable, field, record.get(fieldName));
			}
			Grids grids = new Grids();
			Variables variables = new Variables();
			grids.setGrid(grid);
			variables.setVariable(variable);
			template.setGrids(grids);
			template.setVariables(variables);
			templates.add(template);
		}
		return templates;
	}

	private static void setValueToField(Object object, String field, String value) {
		Field f;
		try {
			f = object.getClass().getDeclaredField(field);
			f.setAccessible(true);
			if (f != null) {
				f.set(object, value);
			}
		} catch (SecurityException e) {
		} catch (NoSuchFieldException e) {
		} catch (IllegalArgumentException e) {
		} catch (IllegalAccessException e) {
		}

	}

	public static List<Map<String, String>> parse(InputStream inStream) {
		List<Map<String, String>> data = new ArrayList<Map<String, String>>();
		try {
			Workbook wb = WorkbookFactory.create(inStream);
			Sheet sheet = wb.getSheetAt(0);
			Row row;
			Cell cell;

			int rows; // No of rows
			rows = sheet.getPhysicalNumberOfRows();

			int cols = 0; // No of columns
			int tmp = 0;

			// This trick ensures that we get the data properly even if it
			// doesn't start from first few rows
			for (int i = 0; i < 10 || i < rows; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					tmp = sheet.getRow(i).getPhysicalNumberOfCells();
					if (tmp > cols)
						cols = tmp;
				}
			}
			List<String> keys = new ArrayList<String>();
			for (int r = 0; r < rows; r++) {
				row = sheet.getRow(r);
				if (row != null) {
					Map<String, String> values = new LinkedHashMap<String, String>();
					for (int c = 0; c < cols; c++) {
						cell = row.getCell((short) c);
						if (r != 0) {
							cell.setCellType(1);
							values.put(keys.get(c), cell.toString());
						}
						if (cell != null) {
							cell.setCellType(1);
							keys.add(cell.toString());
						}
					}
					if (!values.isEmpty()) {
						data.add(values);
					}
				}
			}
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		return data;
	}
}
