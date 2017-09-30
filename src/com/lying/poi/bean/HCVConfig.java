package com.lying.poi.bean;

public class HCVConfig {
	private int sourceSheetNum = 0;
	private int outSheetNum = 1;
	private int sourceCodeNum = 0;
	private int outCodeNum = 6;
	
	private int sourceHcvResultNum = 4;
	private int sourceHcvDateNum = 2;
	private int outHcvResultNum = 7;
	private int outHcvDateNum = 8;
	
	private int outNs5aResultNum = 9;
	private int outNs5aDateNum = 10;
	
	private String sourceFilePath;
	private String outFilePath;
	private String saveFilePath;
	
	private String dateFormat;
	
	public int getSourceSheetNum() {
		return sourceSheetNum;
	}
	public void setSourceSheetNum(int sourceSheetNum) {
		this.sourceSheetNum = sourceSheetNum;
	}
	public int getOutSheetNum() {
		return outSheetNum;
	}
	public void setOutSheetNum(int outSheetNum) {
		this.outSheetNum = outSheetNum;
	}
	public int getSourceCodeNum() {
		return sourceCodeNum;
	}
	public void setSourceCodeNum(int sourceCodeNum) {
		this.sourceCodeNum = sourceCodeNum;
	}
	public int getOutCodeNum() {
		return outCodeNum;
	}
	public void setOutCodeNum(int outCodeNum) {
		this.outCodeNum = outCodeNum;
	}
	public int getSourceHcvResultNum() {
		return sourceHcvResultNum;
	}
	public void setSourceHcvResultNum(int sourceHcvResultNum) {
		this.sourceHcvResultNum = sourceHcvResultNum;
	}
	public int getSourceHcvDateNum() {
		return sourceHcvDateNum;
	}
	public void setSourceHcvDateNum(int sourceHcvDateNum) {
		this.sourceHcvDateNum = sourceHcvDateNum;
	}
	public int getOutHcvResultNum() {
		return outHcvResultNum;
	}
	public void setOutHcvResultNum(int outHcvResultNum) {
		this.outHcvResultNum = outHcvResultNum;
	}
	public int getOutHcvDateNum() {
		return outHcvDateNum;
	}
	public void setOutHcvDateNum(int outHcvDateNum) {
		this.outHcvDateNum = outHcvDateNum;
	}
	public int getOutNs5aResultNum() {
		return outNs5aResultNum;
	}
	public void setOutNs5aResultNum(int outNs5aResultNum) {
		this.outNs5aResultNum = outNs5aResultNum;
	}
	public int getOutNs5aDateNum() {
		return outNs5aDateNum;
	}
	public void setOutNs5aDateNum(int outNs5aDateNum) {
		this.outNs5aDateNum = outNs5aDateNum;
	}
	public String getSourceFilePath() {
		return sourceFilePath;
	}
	public void setSourceFilePath(String sourceFilePath) {
		this.sourceFilePath = sourceFilePath;
	}
	public String getOutFilePath() {
		return outFilePath;
	}
	public void setOutFilePath(String outFilePath) {
		this.outFilePath = outFilePath;
	}
	public String getSaveFilePath() {
		return saveFilePath;
	}
	public void setSaveFilePath(String saveFilePath) {
		this.saveFilePath = saveFilePath;
	}
	
	
	public String getDateFormat() {
		return dateFormat;
	}
	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}
	@Override
	public String toString() {
		return "HCVConfig [sourceSheetNum=" + sourceSheetNum + ", outSheetNum=" + outSheetNum + ", sourceCodeNum="
				+ sourceCodeNum + ", outCodeNum=" + outCodeNum + ", sourceHcvResultNum=" + sourceHcvResultNum
				+ ", sourceHcvDateNum=" + sourceHcvDateNum + ", outHcvResultNum=" + outHcvResultNum + ", outHcvDateNum="
				+ outHcvDateNum + ", outNs5aResultNum=" + outNs5aResultNum + ", outNs5aDateNum=" + outNs5aDateNum
				+ ", sourceFilePath=" + sourceFilePath + ", outFilePath=" + outFilePath + ", saveFilePath="
				+ saveFilePath + ", dateFormat=" + dateFormat + "]";
	}
	
	
}
