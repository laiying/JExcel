package com.lying.poi.ui;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.runtime.IProgressMonitor;
import org.eclipse.jface.action.Action;
import org.eclipse.jface.action.MenuManager;
import org.eclipse.jface.action.StatusLineManager;
import org.eclipse.jface.action.ToolBarManager;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.jface.window.ApplicationWindow;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.List;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;

import com.google.gson.Gson;
import com.google.gson.JsonSyntaxException;
import com.lying.poi.bean.HCVConfig;
import com.lying.poi.utils.DateUtils;
import com.lying.poi.utils.FileUtils;

import org.eclipse.swt.events.SelectionListener;
import java.util.function.Consumer;


public class MainApplication extends ApplicationWindow {

	private static final String[] FILTER_NAMES = {
			"Microsoft Excel Spreadsheet Files (*.xlsx)",
			"All Files (*.*)" };

	// These filter extensions are used to filter which files are displayed.
	private static final String[] FILTER_EXTS = { "*.xlsx", "*.*" };

	 java.util.List<String> filePaths = new ArrayList<>();
	String savePath = null;

	private Action action;

	private List listWidgets;
	private Text savePathTxt;
	private Text sourceSheetIndex;
	private Text outSheetIndex;
	private Text sourceCodeIndex;
	private Text outCodeIndex;
	private Text sourceHcvResultIndex;
	private Text outHcvResultIndex;
	private Text sourceHcvDateIndex;
	private Text outHcvDateIndex;
	private Text outNs5aResultIndex;
	private Text outNs5aDateIndex;
	private Text dateFormatTxt;
	/**
	 * Create the application window.
	 */
	public MainApplication() {
		super(null);
		createActions();
		addToolBar(SWT.FLAT | SWT.WRAP);
		addMenuBar();
		addStatusLine();
	}

	/**
	 * Create contents of the application window.
	 * 
	 * @param parent
	 */
	@Override
	protected Control createContents(Composite parent) {
		Composite container = new Composite(parent, SWT.NONE);
		container.setLayout(new GridLayout(4, false));
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u6570\u636E\u6E90\u7B2C\u51E0\u5F20\u8868:");
		}
		{
			sourceSheetIndex = new Text(container, SWT.BORDER);
			sourceSheetIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u751F\u6210\u6E90\u7B2C\u51E0\u5F20\u8868:");
		}
		{
			outSheetIndex = new Text(container, SWT.BORDER);
			outSheetIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u6570\u636E\u6E90\u6761\u5F62\u7801\u7D22\u5F15:");
		}
		{
			sourceCodeIndex = new Text(container, SWT.BORDER | SWT.WRAP);
			sourceCodeIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u751F\u6210\u6E90\u6761\u5F62\u7801\u7D22\u5F15:");
		}
		{
			outCodeIndex = new Text(container, SWT.BORDER);
			outCodeIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u6570\u636E\u6E90\u5B9E\u9A8C\u7ED3\u679C\u7D22\u5F15:");
		}
		{
			sourceHcvResultIndex = new Text(container, SWT.BORDER);
			sourceHcvResultIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u751F\u6210\u6E90\u5B9E\u9A8C\u7ED3\u679C\u7D22\u5F15:");
		}
		{
			outHcvResultIndex = new Text(container, SWT.BORDER);
			outHcvResultIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u6570\u636E\u6E90\u7ED3\u679C\u65E5\u671F\u7D22\u5F15:");
		}
		{
			sourceHcvDateIndex = new Text(container, SWT.BORDER);
			sourceHcvDateIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u751F\u6210\u6E90\u7ED3\u679C\u65E5\u671F\u7D22\u5F15:");
		}
		{
			outHcvDateIndex = new Text(container, SWT.BORDER);
			outHcvDateIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		{
			Label lblNsa = new Label(container, SWT.NONE);
			lblNsa.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			lblNsa.setText("NS5A\u7ED3\u679C\u7D22\u5F15:");
		}
		{
			outNs5aResultIndex = new Text(container, SWT.BORDER);
			outNs5aResultIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label label = new Label(container, SWT.NONE);
			label.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			label.setText("\u65E5\u671F\u683C\u5F0F:");
		}
		{
			dateFormatTxt = new Text(container, SWT.BORDER);
			dateFormatTxt.setText("yyyy/M/d");
			dateFormatTxt.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		{
			Label lblNsa_1 = new Label(container, SWT.NONE);
			lblNsa_1.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, false, false, 1, 1));
			lblNsa_1.setText("NS5A\u65E5\u671F\u7D22\u5F15:");
		}
		{
			outNs5aDateIndex = new Text(container, SWT.BORDER);
			outNs5aDateIndex.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		{
			Button saveSetting = new Button(container, SWT.NONE);
			saveSetting.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					saveConfig();
				}
			});
			saveSetting.setText("\u4FDD\u5B58\u8BBE\u7F6E");
		}
		{
			Button openFileBtn = new Button(container, SWT.NONE);
			openFileBtn.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					FileDialog dlg = new FileDialog(getShell(), SWT.MULTI);
					dlg.setFilterNames(FILTER_NAMES);
					dlg.setFilterExtensions(FILTER_EXTS);
					String fn = dlg.open();
					if (fn != null) {
						// Append all the selected files. Since getFileNames() returns only
						// the names, and not the path, prepend the path, normalizing
						// if necessary

						String[] files = dlg.getFileNames();
						listWidgets.removeAll();
						filePaths.clear();
						for (int i = 0, n = files.length; i < n; i++) {
							StringBuffer buf = new StringBuffer();
							buf.append(dlg.getFilterPath());
							if (buf.charAt(buf.length() - 1) != File.separatorChar) {
								buf.append(File.separatorChar);
							}
							buf.append(files[i]);
							String filePath = buf.toString();
							listWidgets.add(filePath);
							filePaths.add(filePath);
						}
					}
				}

			});
			openFileBtn.setText("请选择文件");
		}
		new Label(container, SWT.NONE);
		{
			Button saveFileBtn = new Button(container, SWT.NONE);
			saveFileBtn.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					FileDialog dlg = new FileDialog(getShell(), SWT.SAVE);
					String fn = dlg.open();
					if(fn != null) {
						savePath = dlg.getFilterPath()+File.separatorChar+ dlg.getFileName();
						savePathTxt.setText(savePath);
					}
				}

			});
			saveFileBtn.setText("请选择保存路径");
		}
		{
			Button startBtn = new Button(container, SWT.NONE);
			startBtn.addSelectionListener(new SelectionAdapter() {
				@Override
				public void widgetSelected(SelectionEvent e) {
					if(filePaths.size()<2) {
						return;
					}
					
					try {
						int sourceSheetNum = Integer.parseInt(sourceSheetIndex.getText());
						int outSheetNum = Integer.parseInt(outSheetIndex.getText());
						int sourceCodeNum = Integer.parseInt(sourceCodeIndex.getText());
						int outCodeNum = Integer.parseInt(outCodeIndex.getText());
						
						int sourceHcvResultNum = Integer.parseInt(sourceHcvResultIndex.getText());
						int sourceHcvDateNum = Integer.parseInt(sourceHcvDateIndex.getText());
						int outHcvResultNum = Integer.parseInt(outHcvResultIndex.getText());
						int outHcvDateNum = Integer.parseInt(outHcvDateIndex.getText());
						
						int outNs5aResultNum = Integer.parseInt(outNs5aResultIndex.getText());
						int outNs5aDateNum = Integer.parseInt(outNs5aDateIndex.getText());
						
						String dateFormat = dateFormatTxt.getText().trim();
						
						HCVConfig config = new HCVConfig();
						config.setSourceSheetNum(sourceSheetNum);
						config.setOutSheetNum(outSheetNum);
						config.setSourceCodeNum(sourceCodeNum);
						config.setOutCodeNum(outCodeNum);
						
						config.setSourceHcvResultNum(sourceHcvResultNum);
						config.setSourceHcvDateNum(sourceHcvDateNum);
						config.setOutHcvResultNum(outHcvResultNum);
						config.setOutHcvDateNum(outHcvDateNum);
						
						config.setOutNs5aResultNum(outNs5aResultNum);
						config.setOutNs5aDateNum(outNs5aDateNum);
						
						config.setDateFormat(dateFormat);
						
				          new ProgressMonitorDialog(getShell()).run(true, true,
				              new OptRunning(config));
				        } catch (InvocationTargetException e1) {
				        	e1.printStackTrace();
				          MessageDialog.openError(getShell(), "Error", e1.getMessage());
				        } catch (InterruptedException e2) {
				          MessageDialog.openInformation(getShell(), "Cancelled", e2.getMessage());
				        } catch (NumberFormatException e3) {
							// TODO Auto-generated catch block
							e3.printStackTrace();
							 MessageDialog.openInformation(getShell(), "NumberFormatException", e3.getMessage());
						}catch(RuntimeException e4) {
							MessageDialog.openInformation(getShell(), "Error", e4.getMessage());
						}finally {
//							MessageDialog.openInformation(getShell(), "Result", e4.getMessage());
						}
					
				}
			});
			startBtn.setText("\u5F00\u59CB\u8F6C\u6362");
		}
		{
			listWidgets = new List(container, SWT.BORDER | SWT.V_SCROLL);
			GridData gridData = new GridData(GridData.VERTICAL_ALIGN_FILL);
			gridData.widthHint = 284;
			gridData.horizontalSpan = 2;
			gridData.verticalSpan = 5;
			listWidgets.setLayoutData(gridData);
		}
		{
			savePathTxt = new Text(container, SWT.BORDER | SWT.WRAP);
			savePathTxt.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
		}
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		new Label(container, SWT.NONE);
		
		HCVConfig hcvConfig = readConfig();
		if(hcvConfig == null) {
			hcvConfig = new HCVConfig();
		}
		
		sourceSheetIndex.setText(String.valueOf(hcvConfig.getSourceSheetNum()));
		outSheetIndex.setText(String.valueOf(hcvConfig.getOutSheetNum()));
		sourceCodeIndex.setText(String.valueOf(hcvConfig.getSourceCodeNum()));
		outCodeIndex.setText(String.valueOf(hcvConfig.getOutCodeNum()));
		
		sourceHcvResultIndex.setText(String.valueOf(hcvConfig.getSourceHcvResultNum()));
		outHcvResultIndex.setText(String.valueOf(hcvConfig.getOutHcvResultNum()));
		sourceHcvDateIndex.setText(String.valueOf(hcvConfig.getSourceHcvDateNum()));
		outHcvDateIndex.setText(String.valueOf(hcvConfig.getOutHcvDateNum()));
		
		outNs5aResultIndex.setText(String.valueOf(hcvConfig.getOutNs5aResultNum()));
		outNs5aDateIndex.setText(String.valueOf(hcvConfig.getOutNs5aDateNum()));
		
		if(null != hcvConfig.getSourceFilePath()) {
			listWidgets.add(hcvConfig.getSourceFilePath());
			filePaths.add(hcvConfig.getSourceFilePath());
		}
		if(null != hcvConfig.getOutFilePath()) {
			listWidgets.add(hcvConfig.getOutFilePath());
			filePaths.add(hcvConfig.getOutFilePath());
		}
		if(null != hcvConfig.getSaveFilePath()) {
			savePath = hcvConfig.getSaveFilePath();
			savePathTxt.setText(savePath);
		}
		
		if(null == hcvConfig.getDateFormat() || "".equals(hcvConfig.getDateFormat())) {
			dateFormatTxt.setText("yyyy/M/d");
		}else {
			dateFormatTxt.setText(hcvConfig.getDateFormat());
		}

		return container;
	}

	/**
	 * Create the actions.
	 */
	private void createActions() {
		// Create the actions
		{
			action = new Action("File") {

			};
		}
	}

	/**
	 * Create the menu manager.
	 * 
	 * @return the menu manager
	 */
	@Override
	protected MenuManager createMenuManager() {
		MenuManager menuManager = new MenuManager("menu");
		{
			MenuManager menuManager_1 = new MenuManager("New MenuManager");
			menuManager_1.setMenuText("MenuManager");
			menuManager.add(menuManager_1);
		}
		return menuManager;
	}

	/**
	 * Create the toolbar manager.
	 * 
	 * @return the toolbar manager
	 */
	@Override
	protected ToolBarManager createToolBarManager(int style) {
		ToolBarManager toolBarManager = new ToolBarManager(style);
		return toolBarManager;
	}

	/**
	 * Create the status line manager.
	 * 
	 * @return the status line manager
	 */
	@Override
	protected StatusLineManager createStatusLineManager() {
		StatusLineManager statusLineManager = new StatusLineManager();
		return statusLineManager;
	}

	/**
	 * Launch the application.
	 * 
	 * @param args
	 */
	public static void main(String args[]) {
		try {
			MainApplication window = new MainApplication();
			window.setBlockOnOpen(true);
			window.open();
			Display.getCurrent().dispose();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Configure the shell.
	 * 
	 * @param newShell
	 */
	@Override
	protected void configureShell(Shell newShell) {
		super.configureShell(newShell);
		newShell.setText("CWL专用");
		newShell.setImage(new Image(newShell.getDisplay(), "img/ic_launcher.png"));
	}

	/**
	 * Return the initial size of the window.
	 */
	@Override
	protected Point getInitialSize() {
		return new Point(800, 600);
	}

	
	private class OptRunning implements IRunnableWithProgress{
		private HCVConfig config;
		public OptRunning(HCVConfig config) {
			this.config = config;
		}

		@Override
		public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException {
			processExcel(config, filePaths.get(0),filePaths.get(1),savePath == null ? "C:\\BMS HCV项目汇报-新.xlsx" : savePath);
		}
		
	}
	
	private void processExcel(HCVConfig config,String sourcePath,String outPath,String createPath) {
		File sourceFile = new File(sourcePath);
		File outFile = new File(outPath);
		try {
			
			int sourceSheetNum = config.getSourceSheetNum();
			int outSheetNum = config.getOutSheetNum();
			int sourceCodeNum = config.getSourceCodeNum();
			int outCodeNum = config.getOutCodeNum();
			
			int sourceHcvResultNum = config.getSourceHcvResultNum();
			int sourceHcvDateNum = config.getSourceHcvDateNum();
			int outHcvResultNum = config.getOutHcvResultNum();
			int outHcvDateNum = config.getOutHcvDateNum();
			
			int outNs5aResultNum = config.getOutNs5aResultNum();
			int outNs5aDateNum = config.getOutNs5aDateNum();
			
			String dateFormat = config.getDateFormat();
			
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook outWorkbook = new XSSFWorkbook(outFile);

			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(sourceSheetNum);
			XSSFSheet outSheet = outWorkbook.getSheetAt(outSheetNum);
//			XSSFSheet outSheet = outWorkbook.cloneSheet(outSheetNum);

			XSSFCellStyle dateStyle = outWorkbook.createCellStyle();
			dateStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			dateStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			dateStyle.setBorderBottom(BorderStyle.THIN);
			dateStyle.setBorderRight(BorderStyle.THIN);
			dateStyle.setAlignment(HorizontalAlignment.CENTER);
			dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			CreationHelper createHelper = outWorkbook.getCreationHelper();
			dateStyle.setDataFormat(createHelper.createDataFormat().getFormat(dateFormat));
			
			XSSFCellStyle centerStyle = outWorkbook.createCellStyle();
			centerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			centerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			centerStyle.setBorderBottom(BorderStyle.THIN);
			centerStyle.setBorderRight(BorderStyle.THIN);
			centerStyle.setAlignment(HorizontalAlignment.CENTER);
			centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

			Map<String,String> unHCVResultMap = new HashMap<>();

			for (int i = 1; i <= sourceSheet.getLastRowNum(); i++) {
				if (null != sourceSheet.getRow(i)) {
					// one row
					XSSFRow row = sourceSheet.getRow(i);

					if (row != null) {
						String barCode = row.getCell(sourceCodeNum).toString();

						for (int j = 1; j <= outSheet.getLastRowNum(); j++) {
							if (null != outSheet.getRow(j)) {
								// one row
								XSSFRow outRow = outSheet.getRow(j);

								String outBarCode = outRow.getCell(outCodeNum).toString();
								if (barCode.equals(outBarCode)) {
									String result = row.getCell(sourceHcvResultNum).toString();
									String date = row.getCell(sourceHcvDateNum).toString();
									
									if (result.startsWith("HCV")) {
										if(outRow.getCell(outHcvResultNum) == null) {
											outRow.createCell(outHcvResultNum);
										}
										
										outRow.getCell(outHcvResultNum).setCellValue(result);
										outRow.getCell(outHcvResultNum).setCellStyle(centerStyle);
										
										if(outRow.getCell(outHcvDateNum) == null) {
											outRow.createCell(outHcvDateNum);
										}
										outRow.getCell(outHcvDateNum).setCellValue(DateUtils.formatDateString(date, "yyyy-MM-dd HH:mm:ss"));
										outRow.getCell(outHcvDateNum).setCellStyle(dateStyle);
										
										
									}else {
										if(result.startsWith("非1b")) {
											if(outRow.getCell(outNs5aResultNum) == null) {
												outRow.createCell(outNs5aResultNum);
											}
											outRow.getCell(outNs5aResultNum).setCellValue(result);
											outRow.getCell(outNs5aResultNum).setCellStyle(centerStyle);

											if(outRow.getCell(outNs5aDateNum) == null) {
												outRow.createCell(outNs5aDateNum);
											}
											outRow.getCell(outNs5aDateNum).setCellValue(DateUtils.formatDateString(date, "yyyy-MM-dd HH:mm:ss"));
											outRow.getCell(outNs5aDateNum).setCellStyle(dateStyle);
										}else {
											//get barCode result
											String firstResult = unHCVResultMap.get(barCode);
											if(firstResult != null && !firstResult.equals("")) {
												result = firstResult+"、"+ result;
											}
											
											if(outRow.getCell(outNs5aResultNum) == null) {
												outRow.createCell(outNs5aResultNum);
											}
											outRow.getCell(outNs5aResultNum).setCellValue(result);
											outRow.getCell(outNs5aResultNum).setCellStyle(centerStyle);

											if(outRow.getCell(outNs5aDateNum) == null) {
												outRow.createCell(outNs5aDateNum);
											}
											outRow.getCell(outNs5aDateNum).setCellValue(DateUtils.formatDateString(date, "yyyy-MM-dd HH:mm:ss"));
											outRow.getCell(outNs5aDateNum).setCellStyle(dateStyle);
											
											unHCVResultMap.put(barCode, result);
										}
										
									}

								}
							}
						}

					}

				}
			}

			OutputStream os = new FileOutputStream(createPath);
			// out
			outWorkbook.write(os);

			os.close();
			outWorkbook.close();
			sourceWorkbook.close();

		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private HCVConfig readConfig() {
		HCVConfig hcvConfig = null;
		String filePath = FileUtils.getPath()+File.separatorChar+"config"+File.separatorChar+"hcvconfig.json";
		
		String json = FileUtils.readFile(filePath);
		if(json == null || json.equals("")) {
			return null;
		}
		
		Gson gson = new Gson();
		try {
			hcvConfig = gson.fromJson(json, HCVConfig.class);
		} catch (JsonSyntaxException e) {
			e.printStackTrace();
		}
		
		return hcvConfig;
	}
	
	private void saveConfig() {
		int sourceSheetNum = 0;
		int outSheetNum = 1;
		int sourceCodeNum = 0;
		int outCodeNum = 6;
		
		int sourceHcvResultNum = 4;
		int sourceHcvDateNum = 2;
		int outHcvResultNum = 7;
		int outHcvDateNum = 8;
		
		int outNs5aResultNum = 9;
		int outNs5aDateNum = 10;
		String dateFormat = "yyyy/M/d";
		try {
			sourceSheetNum = Integer.parseInt(sourceSheetIndex.getText());
			outSheetNum = Integer.parseInt(outSheetIndex.getText());
			sourceCodeNum = Integer.parseInt(sourceCodeIndex.getText());
			outCodeNum = Integer.parseInt(outCodeIndex.getText());
			sourceHcvResultNum = Integer.parseInt(sourceHcvResultIndex.getText());
			sourceHcvDateNum = Integer.parseInt(sourceHcvDateIndex.getText());
			outHcvResultNum = Integer.parseInt(outHcvResultIndex.getText());
			outHcvDateNum = Integer.parseInt(outHcvDateIndex.getText());
			
			outNs5aResultNum = Integer.parseInt(outNs5aResultIndex.getText());
			outNs5aDateNum = Integer.parseInt(outNs5aDateIndex.getText());
			
			dateFormat = dateFormatTxt.getText().trim();
		}catch (Exception e) {
		}
		
		
		HCVConfig config = new HCVConfig();
		config.setSourceSheetNum(sourceSheetNum);
		config.setOutSheetNum(outSheetNum);
		config.setSourceCodeNum(sourceCodeNum);
		config.setOutCodeNum(outCodeNum);
		
		config.setSourceHcvResultNum(sourceHcvResultNum);
		config.setSourceHcvDateNum(sourceHcvDateNum);
		config.setOutHcvResultNum(outHcvResultNum);
		config.setOutHcvDateNum(outHcvDateNum);
		
		config.setOutNs5aResultNum(outNs5aResultNum);
		config.setOutNs5aDateNum(outNs5aDateNum);
		
		String saveFilePath = savePathTxt.getText().trim();
		config.setSaveFilePath(saveFilePath);
		
		String items[] = listWidgets.getItems();
		if(items != null && items.length>0) {
			config.setSourceFilePath(items[0]);
			if(items.length >1) {
				config.setOutFilePath(items[1]);
			}
		}
		
		config.setDateFormat(dateFormat);
		
		Gson gson = new Gson();
		String json = gson.toJson(config);
		
		String filePathDir = FileUtils.getPath()+File.separatorChar+"config";
		String filePath = filePathDir+File.separatorChar+"hcvconfig.json";
		
		File fileDir = new File(filePathDir);
		File file = new File(filePath);
		FileOutputStream fos;
		try {
			if(!fileDir.exists()) {
				fileDir.mkdirs();
			}
			fos = new FileOutputStream(file);
			fos.write(json.getBytes());
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
}
