package bokesoft.xialj.office.wordtmpl.bean;

import java.util.List;

public class TableUnit {
	
	private int tablePos;
	private String caption;
	private String key;
	private String tableType;
	private List<RowUnit> rowUnitList;
	
	public TableUnit() {
		
	}
	
	public TableUnit(int tablePos,String key,String caption) {
		this.tablePos = tablePos;
		this.caption = caption;
		this.key = key;
	}

	public int getTablePos() {
		return tablePos;
	}
	public void setTablePos(int tablePos) {
		this.tablePos = tablePos;
	}
	public String getCaption() {
		return caption;
	}
	public void setCaption(String caption) {
		this.caption = caption;
	}
	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
	public String getTableType() {
		return tableType;
	}
	public void setTableType(String tableType) {
		this.tableType = tableType;
	}
	public List<RowUnit> getRowUnitList() {
		return rowUnitList;
	}
	public void setRowUnitList(List<RowUnit> rowUnitList) {
		this.rowUnitList = rowUnitList;
	}	

}
