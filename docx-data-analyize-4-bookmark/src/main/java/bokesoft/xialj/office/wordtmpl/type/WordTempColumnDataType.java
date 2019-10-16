package bokesoft.xialj.office.wordtmpl.type;

public enum WordTempColumnDataType {
	BOOLEAN, STRING, LONG, INT, NUMBER, DATE, DATETIME, OPTION;

	public static WordTempColumnDataType getType(String typeStr) {
		switch (typeStr) {
			case "TEXT":
			case "STRING":
			case "SELECT":
			case "COMOBOBOX":
				return STRING;
			case "BOOLEAN":
				return BOOLEAN;
			case "LONG":
				return LONG;
			case "INT":
				return INT;
			case "NUMBER":
				return NUMBER;
			case "DATE":
				return DATE;
			case "DATETIME":
				return DATETIME;
			default:
				return STRING;
		}
	}
}
