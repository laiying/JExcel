package com.lying.poi.ui;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
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
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.dialogs.ProgressMonitorDialog;
import org.eclipse.jface.operation.IRunnableWithProgress;
import org.eclipse.jface.viewers.ListViewer;
import org.eclipse.jface.window.ApplicationWindow;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.RowData;
import org.eclipse.swt.layout.RowLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Control;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.ProgressBar;
import org.eclipse.swt.widgets.Shell;

/**
 * main application window
 * @author laiying
 *
 */
public class MyApplication extends ApplicationWindow{
	
	 // These filter names are displayed to the user in the file dialog. Note that
	  // the inclusion of the actual extension in parentheses is optional, and
	  // doesn't have any effect on which files are displayed.
	  private static final String[] FILTER_NAMES = {
//	      "OpenOffice.org Spreadsheet Files (*.sxc)",
	      "Microsoft Excel Spreadsheet Files (*.xls)",
	      "Microsoft Excel Spreadsheet Files (*.xlsx)",
	      "Comma Separated Values Files (*.csv)", "All Files (*.*)"
	      };

	  // These filter extensions are used to filter which files are displayed.
	  private static final String[] FILTER_EXTS = { "*.xls", "*.xlsx", "*.csv", "*.*"};

	  
	  List<String> filePaths = new ArrayList<>();
	  String savePath = null;
	  
	public MyApplication(Shell parentShell) {
		super(parentShell);
	}

	@Override
	protected void configureShell(Shell shell) {
		shell.setText("cwl");
		shell.setLayout(new RowLayout(SWT.VERTICAL));
		
		shell.setImage(new Image(shell.getDisplay(), "img/ic_launcher.png"));
		
		shell.setSize(800, 600);
	}

	@Override
	protected Control createContents(Composite parent) {
		Button openFile = new Button(parent, SWT.NONE);
		openFile.setText("请选择文件");
		Button saveFile = new Button(parent, SWT.NONE);
		saveFile.setText("请选择保存路径");
		Button startBtn = new Button(parent, SWT.NONE);
		startBtn.setText("开始转换");
		ListViewer listViewer = new ListViewer(parent);
//		listViewer.getControl().setLayoutData(new GridData(GridData.FILL_HORIZONTAL));
		listViewer.getControl().setLayoutData(new RowData(350, 50));
		saveFile.addSelectionListener(new SelectionAdapter() {
			
			@Override
			public void widgetSelected(SelectionEvent e) {
				FileDialog dlg = new FileDialog(getShell(), SWT.SAVE);
				String fn = dlg.open();
				if(fn != null) {
					savePath = dlg.getFilterPath()+File.separatorChar+ dlg.getFileName();
				}
			}
		});
		
		openFile.addSelectionListener(new SelectionAdapter() {
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
		          for (int i = 0, n = files.length; i < n; i++) {
		        	  StringBuffer buf = new StringBuffer();
		            buf.append(dlg.getFilterPath());
		            if (buf.charAt(buf.length() - 1) != File.separatorChar) {
		              buf.append(File.separatorChar);
		            }
		            buf.append(files[i]);
		            String filePath = buf.toString();
		            listViewer.add(filePath);
		            filePaths.add(filePath);
		          }
		        }
			}
		});
		startBtn.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				if(filePaths.size()<2) {
					return;
				}
				try {
			          new ProgressMonitorDialog(getShell()).run(true, true,
			              new OptRunning());
			        } catch (InvocationTargetException e1) {
			          MessageDialog.openError(getShell(), "Error", e1.getMessage());
			        } catch (InterruptedException e2) {
			          MessageDialog.openInformation(getShell(), "Cancelled", e2.getMessage());
			        }
				
			}
		});
		
		return parent;
	}
	
	private class OptRunning implements IRunnableWithProgress{

		@Override
		public void run(IProgressMonitor monitor) throws InvocationTargetException, InterruptedException {
			processExcel(filePaths.get(0),filePaths.get(1),savePath == null ? "C:\\BMS HCV项目汇报-新.xlsx" : savePath);
		}
		
	}
	
	private void processExcel(String sourcePath,String outPath,String createPath) {
		File sourceFile = new File(sourcePath);
		File outFile = new File(outPath);
		try {
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFWorkbook outWorkbook = new XSSFWorkbook(outFile);

			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);
			XSSFSheet outSheet = outWorkbook.getSheetAt(1);
//			XSSFSheet outSheet = outWorkbook.cloneSheet(1);

			XSSFCellStyle cellStyle = outWorkbook.createCellStyle();
			cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setAlignment(HorizontalAlignment.CENTER);
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

			XSSFCellStyle centerStyle = outWorkbook.createCellStyle();
			centerStyle.setAlignment(HorizontalAlignment.CENTER);
			centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

			for (int i = 1; i < sourceSheet.getLastRowNum(); i++) {
				if (null != sourceSheet.getRow(i)) {
					// one row
					XSSFRow row = sourceSheet.getRow(i);

					if (row != null) {
						String barCode = row.getCell(0).toString();

						for (int j = 1; j < outSheet.getLastRowNum(); j++) {
							if (null != outSheet.getRow(j)) {
								// one row
								XSSFRow outRow = outSheet.getRow(j);

								String outBarCode = outRow.getCell(6).toString();
								if (barCode.equals(outBarCode)) {
									String result = row.getCell(4).toString();
									String date = row.getCell(2).toString();
									outRow.getCell(7).setCellValue(result);
									// outRow.GetCell(7).CellStyle = cellStyle;
									outRow.getCell(7).setCellStyle(cellStyle);

									if (row.getCell(2).getCellTypeEnum() == CellType.NUMERIC
											&& DateUtil.isCellDateFormatted(row.getCell(2)))
										outRow.getCell(8).setCellValue(row.getCell(2).getDateCellValue());
									else {
										outRow.getCell(8).setCellValue(row.getCell(2).toString());
									}

									// outRow.GetCell(8).SetCellValue(date);

									if (!result.equals("HCV 1b亚型")) {
										outRow.getCell(9).setCellValue("\\");
										outRow.getCell(9).setCellStyle(centerStyle);

										outRow.getCell(10).setCellValue("\\");
										outRow.getCell(10).setCellStyle(centerStyle);
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
}
