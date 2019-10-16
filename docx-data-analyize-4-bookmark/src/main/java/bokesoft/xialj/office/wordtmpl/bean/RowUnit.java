package bokesoft.xialj.office.wordtmpl.bean;

import java.util.List;

public class RowUnit{
	public final static String TYPE_DATAROW = "DATAROW";
	public final static String TYPE_FIXEDROW = "FIXEDROW";
	
	private String rowType;
	private List<ColumnUnit> collist;

	public List<ColumnUnit> getCollist() {
		return collist;
	}
	public void setCollist(List<ColumnUnit> collist) {
		this.collist = collist;
	}
	public String getRowType() {
		return rowType;
	}
	public void setRowType(String rowType) {
		this.rowType = rowType;
	}
}
