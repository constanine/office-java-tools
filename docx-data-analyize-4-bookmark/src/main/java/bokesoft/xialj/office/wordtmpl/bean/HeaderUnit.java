package bokesoft.xialj.office.wordtmpl.bean;

import java.util.List;

public class HeaderUnit extends ComponentDataUnit {
	/** 对应可能的 下拉选项结合 */
	private List<OptionDataUnit> optionList;
	/** 书签值 */
	protected String bookMark;

	public List<OptionDataUnit> getOptionList() {
		return optionList;
	}

	public void setOptionList(List<OptionDataUnit> optionList) {
		this.optionList = optionList;
	}

	public String getBookMark() {
		return bookMark;
	}

	public void setBookMark(String bookMark) {
		this.bookMark = bookMark;
	}
}
