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

import com.sun.org.apache.bcel.internal.generic.DCONST;
import net.rrm.ehour.config.EhourConfig;
import net.rrm.ehour.config.service.ConfigurationService;
import net.rrm.ehour.data.DateRange;
import net.rrm.ehour.persistence.value.ImageLogo;
import net.rrm.ehour.report.reports.ReportData;
import net.rrm.ehour.report.reports.element.FlatReportElement;
import net.rrm.ehour.ui.common.model.DateModel;
import net.rrm.ehour.ui.common.report.PoiUtil;
import net.rrm.ehour.ui.common.report.Report;
import net.rrm.ehour.ui.common.report.excel.CellFactory;
import net.rrm.ehour.ui.common.report.excel.ExcelWorkbook;
import net.rrm.ehour.ui.common.session.EhourWebSession;
import net.rrm.ehour.ui.common.util.WebUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.wicket.model.IModel;
import org.apache.wicket.model.ResourceModel;
import org.apache.wicket.model.StringResourceModel;
import org.apache.wicket.spring.injection.annot.SpringBean;

import java.util.Date;


public class ExportReportHeader extends AbstractExportReportPart
{
    @SpringBean(name = "configurationService")
    private ConfigurationService configurationService;

    private static final int CONSULTANT_COLUMN = 1;
    private static final int COMPANY_COLUMN = 5;
    private static final int CONSULTANT_SIGNATURE_COLUMN = 12;
    private static final int DG_COLUMN = 19;
    private static final int SECTION_LEADER_COLUMN = 24;
    private static final int DATE_COLUMN = 35;
    private static final int COMPANY_SIGNATURE_COLUMN = 5;
    private static final int HEAD_OF_UNIT_COLUMN = 24;
    private static final int CONSULTANT_ADDRESS_COLUMN = 1;
    private static final int TEL_COLUMN = 5;
    private static final int FRAMEWORK_COLUMN = 10;
    private static final int CONTRACT_COLUMN = 18;
    private static final int END_DATE_COLUMN = 21;
    private static final int NUMBER_OF_DAYS_COLUMN = 28;
    private static final int PROJECT_COLUMN = 35;


    public ExportReportHeader(int cellMargin, Sheet sheet, Report report, ExcelWorkbook workbook)
    {
        super(cellMargin, sheet, report, workbook);
    }

    @Override
    public int createPart(int rowNumber)
    {
        FlatReportElement data = getReportElement();
        rowNumber = addFirstTitleRow(rowNumber);
        rowNumber = addFirstTitleDataRow(rowNumber, data);
        rowNumber = addSecondTitleRow(rowNumber);
        rowNumber = addSecondTitleDataRow(rowNumber, data);
        rowNumber = addThirdTitleRow(rowNumber);
        rowNumber = addThirdTitleDataRow(rowNumber, data);
        rowNumber+=2;

        return rowNumber;
    }

    private FlatReportElement getReportElement(){
        ReportData reportData = getReport().getReportData();
        if(reportData != null && reportData.getReportElements() != null && !reportData.getReportElements().isEmpty()){
            return (FlatReportElement) reportData.getReportElements().get(0);
        }

        return null;
    }

    private CellRangeAddress createCell(Row row, int firstColumn, int lastColumn, String text){
        CellFactory.createCell(row, firstColumn, text, getWorkbook());
        CellRangeAddress cellRangeAddress = new CellRangeAddress(row.getRowNum(),row.getRowNum(),firstColumn,lastColumn);
        getSheet().addMergedRegion(cellRangeAddress);

        return cellRangeAddress;
    }

    private void createTitle(Row row, int firstColumn, int lastColumn, String text){
        CellRangeAddress cellRangeAddress = this.createCell(row, firstColumn, lastColumn, text);
        setTitleBorders(cellRangeAddress);
        row.getCell(firstColumn).setCellStyle(getTitleStyle());
    }

    private void createData(Row row, int firstColumn, int lastColumn, String text){
        CellRangeAddress cellRangeAddress = this.createCell(row, firstColumn, lastColumn, text);
        setDataBorders(cellRangeAddress);
        row.getCell(firstColumn).setCellStyle(getDataStyle());
    }

    private int addFirstTitleRow(int rowNumber)
    {
        Row row = getSheet().createRow(rowNumber++);
        row.setHeight((short)500);

        this.createTitle(row, CONSULTANT_COLUMN, COMPANY_COLUMN-2, "Name and first name of consultant");
        this.createTitle(row, COMPANY_COLUMN, CONSULTANT_SIGNATURE_COLUMN-2, "Name of company");
        this.createTitle(row, CONSULTANT_SIGNATURE_COLUMN, DG_COLUMN-2, "Signature of consultant");
        this.createTitle(row, DG_COLUMN, SECTION_LEADER_COLUMN-2, "DG Unit");
        this.createTitle(row, SECTION_LEADER_COLUMN, DATE_COLUMN-2, "In agreement (operational initiation) Name and Signature");
        this.createTitle(row, DATE_COLUMN, DATE_COLUMN+4, "Date");

        return rowNumber;
    }

    private int addFirstTitleDataRow(int rowNumber, FlatReportElement data)
    {
        Row row = getSheet().createRow(rowNumber++);
        row.setHeight((short)600);

        this.createData(row, CONSULTANT_COLUMN, COMPANY_COLUMN-2, data.getUserLastName() + ", " + data.getUserFirstName());
        this.createData(row, COMPANY_COLUMN, CONSULTANT_SIGNATURE_COLUMN-2, "European Dynamics Consortium");
        this.createData(row, CONSULTANT_SIGNATURE_COLUMN, DG_COLUMN-2, "");
        this.createData(row, DG_COLUMN, SECTION_LEADER_COLUMN-2, data.getCustomerCode() + " - " + data.getCustomerName());
        this.createData(row, SECTION_LEADER_COLUMN, DATE_COLUMN-2, "");
        this.createData(row, DATE_COLUMN, DATE_COLUMN+4, "");

        return rowNumber;
    }

    private int addSecondTitleRow(int rowNumber)
{
    Row row = getSheet().createRow(rowNumber++);
    row.setHeight((short)100);

    row = getSheet().createRow(rowNumber++);
    row.setHeight((short)500);

    this.createTitle(row, COMPANY_SIGNATURE_COLUMN, HEAD_OF_UNIT_COLUMN-2, "Signature of company");
    this.createTitle(row, HEAD_OF_UNIT_COLUMN, HEAD_OF_UNIT_COLUMN+8, "Commission \"Conforme aux faits\"                 Name and Signature");

    return rowNumber;
}

    private int addSecondTitleDataRow(int rowNumber, FlatReportElement data)
    {
        Row row = getSheet().createRow(rowNumber++);
        row.setHeight((short)600);

        this.createData(row, COMPANY_SIGNATURE_COLUMN, HEAD_OF_UNIT_COLUMN-2, "" );
        this.createData(row, HEAD_OF_UNIT_COLUMN, HEAD_OF_UNIT_COLUMN+8, "" );

        return rowNumber;
    }

    private int addThirdTitleRow(int rowNumber)
    {
        Row row = getSheet().createRow(rowNumber++);
        row.setHeight((short)100);

        row = getSheet().createRow(rowNumber++);
        row.setHeight((short)500);

        this.createTitle(row, CONSULTANT_ADDRESS_COLUMN, TEL_COLUMN-2, "Internal address of consultant");
        this.createTitle(row, TEL_COLUMN, FRAMEWORK_COLUMN-2, "Tel.");
        this.createTitle(row, FRAMEWORK_COLUMN, CONTRACT_COLUMN-1, "Framework Contract");
        this.createTitle(row, CONTRACT_COLUMN, END_DATE_COLUMN-1, "Specific Contract");
        this.createTitle(row, END_DATE_COLUMN, NUMBER_OF_DAYS_COLUMN-2, "End date for services in SC");
        this.createTitle(row, NUMBER_OF_DAYS_COLUMN, PROJECT_COLUMN-2, "Number of days of Specific Contract");
        this.createTitle(row, PROJECT_COLUMN, PROJECT_COLUMN+4, "Project");

        return rowNumber;
    }

    private int addThirdTitleDataRow(int rowNumber, FlatReportElement data)
    {
        Row row = getSheet().createRow(rowNumber++);
        row.setHeight((short)600);

        this.createData(row, CONSULTANT_ADDRESS_COLUMN, TEL_COLUMN-2, "" );
        this.createData(row, TEL_COLUMN, FRAMEWORK_COLUMN+-2, "" );
        this.createData(row, FRAMEWORK_COLUMN, CONTRACT_COLUMN-1, "" );
        this.createData(row, CONTRACT_COLUMN, END_DATE_COLUMN-1, "" );
        this.createData(row, END_DATE_COLUMN, NUMBER_OF_DAYS_COLUMN-2, "" );
        this.createData(row, NUMBER_OF_DAYS_COLUMN, PROJECT_COLUMN-2, "" );
        this.createData(row, PROJECT_COLUMN, PROJECT_COLUMN+4, data.getProjectName() );

        return rowNumber;
    }

    private IModel<String> getExcelReportName(DateRange dateRange)
    {
        EhourConfig config = EhourWebSession.getEhourConfig();

        return new StringResourceModel("excelMonth.reportName",
                null,
                new Object[]{EhourWebSession.getUser().getFullName(),
                             new DateModel(dateRange.getDateStart(), config, DateModel.DATESTYLE_MONTHONLY)});
    }

    private ConfigurationService getConfigurationService()
    {
        if (configurationService == null)
        {
            WebUtils.springInjection(this);
        }

        return configurationService;
    }
}
