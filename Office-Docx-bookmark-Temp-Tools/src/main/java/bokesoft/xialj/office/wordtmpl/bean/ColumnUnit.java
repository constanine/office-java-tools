package bokesoft.xialj.office.wordtmpl.bean;

/**
 * 数据元素
 */
public class ColumnUnit {
	/** 显示标题  */
	protected String caption;
	/** 数据标识key */
	protected String key;
	/** 显示类型 */
	protected String type;
	/** 记录的填充值 */
	protected String record;
	
	public String getRecord() {
		return record;
	}
	public void setRecord(String record) {
		this.record = record;
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
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
}
