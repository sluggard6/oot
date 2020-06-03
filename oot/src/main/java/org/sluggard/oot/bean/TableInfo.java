package org.sluggard.oot.bean;


public class TableInfo {
	
	private String tableName;
	private String tComments;
	private String columnName;
	private String dataType;
	private String dataLength;
	private String dataScale;
	private String nullable;
	private String dataDefault;
	private String cComments;
	public String getTableName() {
		return tableName;
	}
	public void setTableName(String tableName) {
		this.tableName = tableName;
	}
	public String gettComments() {
		return tComments;
	}
	public void settComments(String tComments) {
		this.tComments = tComments;
	}
	public String getColumnName() {
		return columnName;
	}
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}
	public String getDataType() {
		return dataType;
	}
	public void setDataType(String dataType) {
		this.dataType = dataType;
	}
	public String getDataLength() {
		return dataLength;
	}
	public void setDataLength(String dataLength) {
		this.dataLength = dataLength;
	}
	public String getDataScale() {
		return dataScale;
	}
	public void setDataScale(String dataScale) {
		this.dataScale = dataScale;
	}
	public String getNullable() {
		return nullable;
	}
	public void setNullable(String nullable) {
		this.nullable = nullable;
	}
	public String getDataDefault() {
		return dataDefault;
	}
	public void setDataDefault(String dataDefault) {
		this.dataDefault = dataDefault;
	}
	public String getcComments() {
		return cComments;
	}
	public void setcComments(String cComments) {
		this.cComments = cComments;
	}
	@Override
	public String toString() {
		return "TableInfo [tableName=" + tableName + ", tComments=" + tComments + ", columnName=" + columnName
				+ ", dataType=" + dataType + ", dataLength=" + dataLength + ", dataScale=" + dataScale + ", nullable="
				+ nullable + ", dataDefault=" + dataDefault + ", cComments=" + cComments + "]";
	}

}
