package com.ruite.util.poi;

/**
 * 数据，保存坐标与具体内容
 * @author fangzhiyang
 *
 */
public class MyData {

	private String text;
	private int row;
	private int cell;

	public MyData(String text, int row, int cell) {
		this.text = text;
		this.row = row;
		this.cell = cell;
	}
	
	public MyData(int row, int cell) {
		this.row = row;
		this.cell = cell;
	}
	
	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getCell() {
		return cell;
	}

	public void setCell(int cell) {
		this.cell = cell;
	}

}
