package bokesoft.xialj.office.wordtmpl.type;

public enum WordTempComponentType {
	TEXT,OPTION,COMOBOBOX,SHOW,BOOLEAN,LONG,DATE,NUMBER,DATETIME;
	
	public static WordTempComponentType getType(String typeStr) {
		switch (typeStr) {
		case "TEXT":
		case "STRING":
			return TEXT;
		case "COMOBOBOX":
		case "SELECT":
			return COMOBOBOX;
		case "BOOLEAN":
			return BOOLEAN;
		case "SHOW":
			return SHOW;
		case "LONG":
		case "INT":
			return LONG;
		case "NUMBER":
			return NUMBER;
		case "DATE":
			return DATE;
		case "DATETIME":
			return DATETIME;
		default:
			return TEXT;
		}
	}
}
