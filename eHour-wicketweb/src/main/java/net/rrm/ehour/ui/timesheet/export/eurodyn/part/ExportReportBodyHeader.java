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

import net.rrm.ehour.ui.common.report.Report;
import net.rrm.ehour.ui.common.report.excel.CellFactory;
import net.rrm.ehour.ui.common.report.excel.ExcelStyle;
import net.rrm.ehour.ui.common.report.excel.ExcelWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.wicket.model.ResourceModel;

/**
 * Created on Mar 25, 2009, 7:16:55 AM
 * @author Thies Edeling (thies@te-con.nl) 
 *
 */
public class ExportReportBodyHeader extends AbstractExportReportPart
{

	private int maxColumns = 0;
	public ExportReportBodyHeader(int cellMargin, Sheet sheet, Report report, ExcelWorkbook workbook)
	{
		super(cellMargin, sheet, report, workbook);
	}

	private void createCell(Row row, int column, String text){
		CellFactory.createCell(row, column, text, getWorkbook());
		row.getCell(column).setCellStyle(getTitleStyle());
	}

	private void createDataCell(Row row, int column, String text){
		CellFactory.createCell(row, column, text, getWorkbook());
		row.getCell(column).setCellStyle(getDataStyle());
	}
	
	public int createPart(int rowNumber)
	{
		rowNumber = this.createHeaderColumns(rowNumber);
		rowNumber = this.createMonthLines(rowNumber);
		return rowNumber;
	}

	private int createHeaderColumns(int rowNumber){
		Row row = getSheet().createRow(rowNumber);
		row.setHeight((short)800);
		String[] headers = {
				"MONTH", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17",
				"18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31",
				"Days worked", "Days remaining", "Estimate by consultant of end of services"
		};
		int i =1;
		for(String header:headers){
			this.createCell(row, i++, header);
		}
		maxColumns = i;

		return ++rowNumber;
	}

	private int createMonthLines(int rowNumber){
		String[] months = {
				"JANUARY",
				"FEBRUARY",
				"MARCH",
				"APRIL",
				"MAY",
				"JUNE",
				"JULY",
				"AUGUST",
				"SEPTEMBER",
				"OCTOBER",
				"NOVEMBER",
				"DECEMBER"
		};
		for(String month:months){
			int column = 1;
		    Row row = getSheet().createRow(rowNumber++);
			this.createCell(row, column++, month);

			//Empty cells for the days of the month
            while(column <= 33){
                this.createDataCell(row, column++, "");
            }
			this.createDataCell(row, column++, "120");
			this.createDataCell(row, column++, "");
		}

		return rowNumber;
	}
}
