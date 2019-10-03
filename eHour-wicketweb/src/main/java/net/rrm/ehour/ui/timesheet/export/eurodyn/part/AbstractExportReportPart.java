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
import net.rrm.ehour.report.reports.element.FlatReportElement;
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
import java.util.Map;
import java.util.TreeMap;

/**
 * Created on Mar 25, 2009, 3:34:34 PM
 * @author Thies Edeling (thies@te-con.nl) 
 *
 */
public abstract class AbstractExportReportPart
{
    protected static final String CONTRACTOR_CUSTOMER = "Contractor";

	private int cellMargin;
	private EhourConfig config;
	private SimpleDateFormat formatter;
	private Sheet sheet;
	private Report report;
	private ExcelWorkbook workbook;

	private CellStyle titleStyle;
	private CellStyle dataStyle;
	private CellStyle weekendStyle;
	private CellStyle holidayStyle;
	private CellStyle footerStyle;
	private CellStyle footerBorderStyle;

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

		weekendStyle = workbook.getWorkbook().createCellStyle();
		weekendStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		weekendStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		weekendStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		weekendStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		weekendStyle.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);
		weekendStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		weekendStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		holidayStyle = workbook.getWorkbook().createCellStyle();
		holidayStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		holidayStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		holidayStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		holidayStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		holidayStyle.setVerticalAlignment(CellStyle.VERTICAL_JUSTIFY);
		holidayStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.index);
		holidayStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		footerStyle = workbook.getWorkbook().createCellStyle();
		Font footerFont = workbook.getWorkbook().createFont();
		footerFont.setFontName("Arial");
		footerFont.setFontHeightInPoints((short)8);
		footerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		footerStyle.setFont(footerFont);

		footerBorderStyle = workbook.getWorkbook().createCellStyle();
		footerBorderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		footerBorderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		footerBorderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		footerBorderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		footerBorderStyle.setFont(footerFont);
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

	protected boolean isContractorElement(FlatReportElement element){
	    return element.getCustomerCode().equals(CONTRACTOR_CUSTOMER);
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

	protected CellStyle getWeekendStyle() {
		return weekendStyle;
	}

	protected CellStyle getHolidayStyle() {
		return holidayStyle;
	}

	protected CellStyle getFooterStyle() {
		return footerStyle;
	}

	protected CellStyle getFooterBorderStyle() {
		return footerBorderStyle;
	}

	protected void createEmptyCells(Row row, ExcelStyle excelStyle)
	{

	}	
}
