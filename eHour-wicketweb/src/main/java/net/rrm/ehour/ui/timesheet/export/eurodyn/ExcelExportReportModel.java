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

import net.rrm.ehour.domain.Project;
import net.rrm.ehour.domain.ProjectAssignment;
import net.rrm.ehour.report.criteria.ReportCriteria;
import net.rrm.ehour.report.reports.ReportData;
import net.rrm.ehour.report.reports.element.FlatReportElement;
import net.rrm.ehour.report.reports.element.ReportElement;
import net.rrm.ehour.report.service.DetailedReportService;
import net.rrm.ehour.sort.ProjectAssignmentComparator;
import net.rrm.ehour.ui.common.report.AbstractReportModel;
import net.rrm.ehour.ui.common.util.WebUtils;
import net.rrm.ehour.util.DateUtil;
import org.apache.log4j.Logger;
import org.apache.wicket.spring.injection.annot.SpringBean;

import java.text.ParseException;
import java.util.*;

import static org.springframework.util.Assert.notNull;

/**
 * Print report for printing a timesheet
 */

public class ExcelExportReportModel extends AbstractReportModel {
    private static final long serialVersionUID = -8062083697181324496L;
    private static final Logger LOGGER = Logger.getLogger(ExcelExportReportModel.class);

    @SpringBean
    private DetailedReportService detailedReportService;

    private transient SortedMap<ProjectAssignment, Map<Date, FlatReportElement>> rowMap;

    public ExcelExportReportModel(ReportCriteria criteria) {
        super(criteria);
    }

    protected ProjectAssignment getRowKey(FlatReportElement aggregate) {
        ProjectAssignment pa = new ProjectAssignment();

        pa.setAssignmentId(aggregate.getAssignmentId());
        pa.setRole(aggregate.getRole());

        Project prj = new Project();
        prj.setName(aggregate.getProjectName());
        prj.setProjectId(aggregate.getProjectId());

        pa.setProject(prj);

        return pa;
    }

    /**
     * Format is ddmmyyyy
     *
     * @throws ParseException
     */
    protected Date getAggregateDate(FlatReportElement aggregate) throws ParseException {
        return aggregate.getDayDate();
    }

    protected Comparator<ProjectAssignment> getRKComparator() {
        return new ProjectAssignmentComparator();
    }

    protected ReportData fetchReportData(ReportCriteria reportCriteria) {
        return getDetailedReportService().getDetailedReportData(reportCriteria);
    }

    private DetailedReportService getDetailedReportService() {
        if (detailedReportService == null) {
            WebUtils.springInjection(this);
        }

        return detailedReportService;
    }

    @Override
    protected ReportData getReportData(ReportCriteria reportCriteria) {
        Map<Date, FlatReportElement> rowAggregates;
        Date aggregateDate;
        ProjectAssignment rowKey;

        ReportData aggregateData = getValidReportData(reportCriteria);

        rowMap = new TreeMap<>(getRKComparator());

        for (ReportElement element : aggregateData.getReportElements()) {
            FlatReportElement aggregate = (FlatReportElement) element;

            rowKey = getRowKey(aggregate);

            if (rowMap.containsKey(rowKey)) {
                rowAggregates = rowMap.get(rowKey);
            } else {
                rowAggregates = new HashMap<>();
            }

            aggregateDate = getValidAggregateDate(aggregate);

            if(aggregateDate != null){
                aggregateDate = DateUtil.nullifyTime(aggregateDate);
                rowAggregates.put(aggregateDate, aggregate);

                rowMap.put(rowKey, rowAggregates);
            }
        }

        return aggregateData;
    }

    private ReportData getValidReportData(ReportCriteria reportCriteria) {
        ReportData reportData = fetchReportData(reportCriteria);

        notNull(reportData);
        notNull(reportData.getReportElements());

        return reportData;
    }

    /**
     * Get grand total hours
     *
     * @return
     */
    public float getGrandTotalHours() {
        float total = 0;
        Map<Date, FlatReportElement> aggMap;

        for (ProjectAssignment key : rowMap.keySet()) {
            aggMap = rowMap.get(key);

            for (Map.Entry<Date, FlatReportElement> entry : aggMap.entrySet()) {
                Number n = entry.getValue().getTotalHours();

                if (n != null) {
                    total += n.floatValue();
                }
            }
        }

        return total;
    }

    /**
     * Get the values
     *
     * @return
     */
    public SortedMap<ProjectAssignment, Map<Date, FlatReportElement>> getValues() {
        if (rowMap == null) {
            lazyInitRowMap();
        }
        return rowMap;
    }

    private void lazyInitRowMap() {
        load();
    }

    private Date getValidAggregateDate(FlatReportElement aggregate) {
        Date date;

        try {
            date = getAggregateDate(aggregate);
        } catch (ParseException e) {
            LOGGER.warn("failed to parse date of " + aggregate, e);
            date = new Date();
        }

        return date;
    }

    /*
          * (non-Javadoc)
          * @see org.apache.wicket.model.LoadableDetachableModel#onDetach()
          */
    @Override
    protected void onDetach() {
        rowMap = null;
    }
}
