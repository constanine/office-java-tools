package bokesoft.xialj.office.wordtmpl.type;

public enum WordTempColumnControllerType {
	TEXT,OPTION,COMOBOBOX,SHOW,LONG,DATE,NUMBER;
	
	public static WordTempColumnControllerType getType(String typeStr) {
		switch (typeStr) {
		case "TEXT":
		case "STRING":
			return TEXT;
		case "COMOBOBOX":
		case "SELECT":
			return COMOBOBOX;
		case "SHOW":
			return SHOW;
		case "LONG":
		case "INT":
			return LONG;
		case "NUMBER":
			return NUMBER;
		case "DATE":
		case "DATETIME":
			return DATE;
		default:
			return TEXT;
		}
	}
}
