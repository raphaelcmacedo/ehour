/*
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 */

package net.rrm.ehour.ui.timesheet.export.eurodyn.part;

import net.rrm.ehour.config.EhourConfig;
import net.rrm.ehour.ui.common.report.Report;
import net.rrm.ehour.ui.common.report.excel.CellFactory;
import net.rrm.ehour.ui.common.report.excel.ExcelStyle;
import net.rrm.ehour.ui.common.report.excel.ExcelWorkbook;
import net.rrm.ehour.ui.common.session.EhourWebSession;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.text.SimpleDateFormat;
import java.util.Locale;

/**
 * Created on Mar 25, 2009, 3:34:34 PM
 * @author Thies Edeling (thies@te-con.nl) 
 *
 */
public abstract class AbstractExportReportPart
{
	private int cellMargin;
	private EhourConfig config;
	private SimpleDateFormat formatter;
	private Sheet sheet;
	private Report report;
	private ExcelWorkbook workbook;

	private CellStyle titleStyle;
	private CellStyle dataStyle;

	public AbstractExportReportPart(int cellMargin, Sheet sheet, Report report, ExcelWorkbook workbook)
	{
		this.cellMargin = cellMargin;
		this.sheet = sheet;
		this.report = report;
		this.workbook = workbook;
		
		init();
	}
	
	public abstract int createPart(int rowNumber);
	
	private void init()
	{
		config = EhourWebSession.getEhourConfig();
		Locale locale = config.getFormattingLocale();
		formatter = new SimpleDateFormat("dd MMM yy", locale);
		setStyles();
	}

	private void setStyles(){
		titleStyle = workbook.getWorkbook().createCellStyle();
		titleStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
		titleStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
		titleStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
		titleStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
		titleStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);

		Font cellFontTtile = workbook.getWorkbook().createFont();
		cellFontTtile.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cellFontTtile.setFontName("Arial");
		cellFontTtile.setFontHeightInPoints((short)10);
		titleStyle.setFont(cellFontTtile);


		dataStyle = workbook.getWorkbook().createCellStyle();
		dataStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		dataStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		dataStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		dataStyle.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);

		Font cellFontData = workbook.getWorkbook().createFont();
		cellFontData.setFontName("Arial");
		cellFontData.setFontHeightInPoints((short)9);
		dataStyle.setFont(cellFontData);
	}

	protected void setTitleBorders(CellRangeAddress cellRangeAddress){
		RegionUtil.setBorderTop(CellStyle.BORDER_MEDIUM, cellRangeAddress, sheet, workbook.getWorkbook());
		RegionUtil.setBorderBottom(CellStyle.BORDER_MEDIUM, cellRangeAddress, sheet, workbook.getWorkbook());
		RegionUtil.setBorderLeft(CellStyle.BORDER_MEDIUM, cellRangeAddress, sheet, workbook.getWorkbook());
		RegionUtil.setBorderRight(CellStyle.BORDER_MEDIUM, cellRangeAddress, sheet, workbook.getWorkbook());
	}

	protected void setDataBorders(CellRangeAddress cellRangeAddress){
		RegionUtil.setBorderTop(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook.getWorkbook());
		RegionUtil.setBorderBottom(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook.getWorkbook());
		RegionUtil.setBorderLeft(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook.getWorkbook());
		RegionUtil.setBorderRight(CellStyle.BORDER_THIN, cellRangeAddress, sheet, workbook.getWorkbook());
	}

	protected int getCellMargin()
	{
		return cellMargin;
	}

	protected EhourConfig getConfig()
	{
		return config;
	}

	protected SimpleDateFormat getFormatter()
	{
		return formatter;
	}

	protected Sheet getSheet()
	{
		return sheet;
	}

	protected Report getReport()
	{
		return report;
	}

	protected ExcelWorkbook getWorkbook()
	{
		return workbook;
	}

	protected CellStyle getTitleStyle() {
		return titleStyle;
	}

	protected CellStyle getDataStyle() {
		return dataStyle;
	}

	protected void createEmptyCells(Row row, ExcelStyle excelStyle)
	{
		for (int i : ExportReportColumn.EMPTY.getColumns())
		{
			CellFactory.createCell(row, getCellMargin() + i, getWorkbook(), excelStyle);
		}
	}	
}
