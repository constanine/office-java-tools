package bokesoft.xialj.office.wordtmpl.bean;

import java.util.List;

/** 表格中的内容 */
public class TableUnit{
	
	private List<RowUnit> rowlist;
	/** 按表格顺序自动命名 */
	private String key;
	public List<RowUnit> getRowlist() {
		return rowlist;
	}
	public void setRowlist(List<RowUnit> rowlist) {
		this.rowlist = rowlist;
	}
	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
}
