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

package net.rrm.ehour.ui.timesheet.export.eurodyn;

import net.rrm.ehour.report.criteria.ReportCriteria;
import net.rrm.ehour.ui.common.report.ExcelReport;
import net.rrm.ehour.ui.common.report.Report;
import net.rrm.ehour.ui.common.report.excel.ExcelWorkbook;
import net.rrm.ehour.ui.common.util.WebUtils;
import net.rrm.ehour.ui.timesheet.export.TimesheetExportParameter;
import net.rrm.ehour.ui.timesheet.export.eurodyn.part.*;
import org.apache.commons.collections.bidimap.TreeBidiMap;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;
import java.util.TreeMap;

/**
 * Created on Mar 23, 2009, 1:30:04 PM
 *
 * @author Thies Edeling (thies@te-con.nl)
 */
public class TimesheetExcelExport implements ExcelReport {
    private static final long serialVersionUID = -4841781257347819473L;

    private static final int CELL_BORDER = 1;
    private final ReportCriteria reportCriteria;

    public TimesheetExcelExport(ReportCriteria reportCriteria) {
        this.reportCriteria = reportCriteria;
    }

    @Override
    public void write(OutputStream stream) throws IOException {
        ExcelExportReportModel report = new ExcelExportReportModel(reportCriteria);
        ExcelWorkbook workbook = createWorkbook(report);

        workbook.write(stream);
    }

    private ExcelWorkbook createWorkbook(Report report) {
        ExcelWorkbook workbook = new ExcelWorkbook();

        String sheetName = WebUtils.formatDate("yyyy", report.getReportRange().getDateStart());
        Sheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(sheetName));
        sheet.getPrintSetup().setLandscape(true);
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);

        sheet.setColumnWidth(0, 333);
        sheet.setColumnWidth(1, 3300);
        for(int i = 2; i < 33; i++){
            sheet.setColumnWidth(i, 730);
        }

        sheet.setColumnWidth(33, 2100);
        sheet.setColumnWidth(34, 2300);
        sheet.setColumnWidth(35, 3100);

        int rowNumber = 1;

        Map<String, String> comments = new TreeMap<String, String>();
        rowNumber = new ExportReportHeader(CELL_BORDER, sheet, report, workbook).createPart(rowNumber);
        rowNumber = new ExportReportBodyHeader(CELL_BORDER, sheet, report, workbook).createPart(rowNumber);
        rowNumber = new ExportReportBody(CELL_BORDER, sheet, report, workbook, comments).createPart(rowNumber);
        rowNumber = new ExportReportTotal(CELL_BORDER, sheet, report, workbook, comments).createPart(rowNumber);

        return workbook;
    }

    private boolean isInclSignOff(Report report) {
        String key = TimesheetExportParameter.INCL_SIGN_OFF.name();
        Object object = report.getReportCriteria().getUserSelectedCriteria().getCustomParameters().get(key);
        return (object != null) && (Boolean) object;
    }

    @Override
    public String getFilenameWihoutSuffix() {
        return "month_report";
    }
}
