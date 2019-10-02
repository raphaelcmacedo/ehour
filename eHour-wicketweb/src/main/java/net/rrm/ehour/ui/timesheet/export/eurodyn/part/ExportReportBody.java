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

import net.rrm.ehour.report.reports.ReportData;
import net.rrm.ehour.report.reports.element.FlatReportElement;
import net.rrm.ehour.report.reports.element.ReportElement;
import net.rrm.ehour.ui.common.report.Report;
import net.rrm.ehour.ui.common.report.excel.CellFactory;
import net.rrm.ehour.ui.common.report.excel.ExcelStyle;
import net.rrm.ehour.ui.common.report.excel.ExcelWorkbook;
import net.rrm.ehour.util.DateUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * Created on Mar 25, 2009, 6:35:04 AM
 *
 * @author Thies Edeling (thies@te-con.nl)
 */
public class ExportReportBody extends AbstractExportReportPart {

    private static final int ROW_FIRST_MONTH = 12;
    private static final int COLUMN_FIRST_DAY = 1;

    private Map<Integer, Double> daysWorked = new TreeMap<Integer, Double>();
    private Double daysRemaining;

    private void initDaysWorked(){
        for(int i =0; i < 12; i++){
            daysWorked.put(i,0.0);
        }
    }

    public ExportReportBody(int cellMargin, Sheet sheet, Report report, ExcelWorkbook workbook) {
        super(cellMargin, sheet, report, workbook);
        initDaysWorked();
    }

    private void createCell(Calendar calendar, String text, CellStyle cellStyle){
        int rowNumber = calendar.get(Calendar.MONTH) + ROW_FIRST_MONTH;
        int colNumber = calendar.get(Calendar.DAY_OF_MONTH) + COLUMN_FIRST_DAY;

        Row row = getSheet().getRow(rowNumber);

        String currentText = row.getCell(colNumber).getStringCellValue();
        if(NumberUtils.isNumber(currentText)){//If the cell already contains numbers don't replace it
            return;
        }

        CellFactory.createCell(row, colNumber, text, getWorkbook());
        row.getCell(colNumber).setCellStyle(cellStyle);
    }

    @Override
    public int createPart(int rowNumber) {
        markAllWeekends();

        List<FlatReportElement> elements = (List<FlatReportElement>) getReport().getReportData().getReportElements();
        List<Date> dateSequence = DateUtil.createDateSequence(getReport().getReportRange(), getConfig());
        setDataToReport(elements);

        writeSubtotals();

        return rowNumber;
    }

    private void setDataToReport(List<FlatReportElement> elements){
        for (FlatReportElement element : elements){
            if(!isContractorElement(element) && daysRemaining == null){
                daysRemaining = element.getAssignmentDaysAllotted();
            }

            Calendar calendar = Calendar.getInstance();
            calendar.setTime(element.getDayDate());

            String text = "";
            CellStyle cellStyle = getDataStyle();
            if(this.isWeekend(calendar)){
                cellStyle = getWeekendStyle();
            }else if(element.getLocked() != null && element.getLocked()){
                cellStyle = getHolidayStyle();
            }else{
                text = getTextForReport(element, calendar);
            }

            this.createCell(calendar, text, cellStyle);
        }
    }

    private String getTextForReport(FlatReportElement element, Calendar calendar){
        double totalHours = (double) element.getTotalHours();
        int month = calendar.get(Calendar.MONTH);
        double sum = daysWorked.get(month);

        if(isContractorElement(element)){
            return element.getProjectCode();
        }else if(totalHours == 8){
            sum += 1;
            daysWorked.put(month, sum);
            return "1";
        }else if(totalHours > 0 && totalHours < 8){
            sum += 0.5;
            daysWorked.put(month, sum);
            return "0.5";
        }
        return "";
    }

    private boolean isWeekend(Calendar calendar){
        int dayOfWeek = calendar.get(Calendar.DAY_OF_WEEK);
        return (Calendar.SUNDAY == dayOfWeek || Calendar.SATURDAY == dayOfWeek);
    }

    private void markAllWeekends(){
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(new Date());
        calendar.set(Calendar.DAY_OF_MONTH,1);
        calendar.set(Calendar.MONTH,0);

        int year = calendar.get(Calendar.YEAR);
        while(calendar.get(Calendar.YEAR) == year){
            if(isWeekend(calendar)){
                this.createCell(calendar, "", getWeekendStyle());
            }
            calendar.add(Calendar.DATE,1);
        }
    }

    private void writeSubtotals(){
        int colNumberDaysWorked = 32 + COLUMN_FIRST_DAY;
        int colNumberDaysRemaining = 33 + COLUMN_FIRST_DAY;

        for(Integer month : daysWorked.keySet()){
            int rowNumber = month + ROW_FIRST_MONTH;
            Row row = getSheet().getRow(rowNumber);

            double days = daysWorked.get(month);
            daysRemaining -= days;

            CellFactory.createCell(row, colNumberDaysWorked, String.valueOf(days), getWorkbook());
            row.getCell(colNumberDaysWorked).setCellStyle(getDataStyle());

            CellFactory.createCell(row, colNumberDaysRemaining, String.valueOf(daysRemaining), getWorkbook());
            row.getCell(colNumberDaysRemaining).setCellStyle(getDataStyle());


        }
    }



}
