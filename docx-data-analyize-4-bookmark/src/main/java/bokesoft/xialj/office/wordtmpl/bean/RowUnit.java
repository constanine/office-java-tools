package bokesoft.xialj.office.wordtmpl.bean;

import java.util.List;

public class RowUnit {	
	private String rowType;
	private int rowIdx; 
	private List<ColumnUnit> collist;

	public List<ColumnUnit> getCollist() {
		return collist;
	}

	public String getRowType() {
		return rowType;
	}	

	public void setRowType(String rowType) {
		this.rowType = rowType;
	}

	public void setCollist(List<ColumnUnit> collist) {
		this.collist = collist;
	}

	public int getRowIdx() {
		return rowIdx;
	}

	public void setRowIdx(int rowIdx) {
		this.rowIdx = rowIdx;
	}
	
}
