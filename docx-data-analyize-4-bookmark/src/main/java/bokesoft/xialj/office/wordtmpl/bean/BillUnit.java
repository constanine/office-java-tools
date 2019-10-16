package bokesoft.xialj.office.wordtmpl.bean;

import java.util.List;

/**
 * 单据元素
 */
public class BillUnit{
	/**
	 * 头表元素
	 */
	private List<HeaderUnit> headers;
	/**
	 * 明细表元素
	 */
	private List<TableUnit> tables;
	
	public List<HeaderUnit> getHeaders() {
		return headers;
	}
	public void setHeaders(List<HeaderUnit> headers) {
		this.headers = headers;
	}
	public List<TableUnit> getTables() {
		return tables;
	}
	public void setTables(List<TableUnit> tables) {
		this.tables = tables;
	}
}
