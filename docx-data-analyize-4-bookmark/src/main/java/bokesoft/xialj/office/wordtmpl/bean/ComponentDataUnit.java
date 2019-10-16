package bokesoft.xialj.office.wordtmpl.bean;

import bokesoft.xialj.office.wordtmpl.type.WordTempComponentType;
import bokesoft.xialj.office.wordtmpl.type.WordTempDataType;

/**
 * 数据元素
 */
public class ComponentDataUnit {
	/** 显示标题  */
	protected String caption;
	/** 数据标识key */
	protected String key;
	/** 显示类型 */
	protected WordTempComponentType componentType;
	/** 数据类别 */
	protected WordTempDataType dataType;
	/** 记录的填充值 */
	protected String record;
	
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
	public String getRecord() {
		return record;
	}
	public void setRecord(String record) {
		this.record = record;
	}
	public WordTempComponentType getComponentType() {
		return componentType;
	}
	public void setComponentType(WordTempComponentType controllerType) {
		this.componentType = controllerType;
	}
	public WordTempDataType getDataType() {
		return dataType;
	}
	public void setDataType(WordTempDataType dataType) {
		this.dataType = dataType;
	}
	
	
}