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

package net.rrm.ehour.report.service;

import com.google.common.collect.Lists;
import net.rrm.ehour.data.DateRange;
import net.rrm.ehour.domain.Project;
import net.rrm.ehour.domain.ProjectAssignment;
import net.rrm.ehour.domain.User;
import net.rrm.ehour.persistence.project.dao.ProjectAssignmentDao;
import net.rrm.ehour.persistence.project.dao.ProjectDao;
import net.rrm.ehour.persistence.report.dao.DetailedReportDao;
import net.rrm.ehour.persistence.report.dao.ReportAggregatedDao;
import net.rrm.ehour.report.criteria.ReportCriteria;
import net.rrm.ehour.report.reports.ReportData;
import net.rrm.ehour.report.reports.element.FlatReportElement;
import net.rrm.ehour.report.reports.element.FlatReportElementBuilder;
import net.rrm.ehour.report.reports.element.LockableDate;
import net.rrm.ehour.timesheet.service.TimesheetLockService;
import net.rrm.ehour.util.DomainUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.Date;
import java.util.List;

/**
 * Report service for detailed reports implementation
 */
@Service("detailedReportService")
public class DetailedReportServiceImpl extends AbstractReportServiceImpl<FlatReportElement> implements DetailedReportService {
    private static final String CONTRACTOR_CUSTOMER_CODE = "Contractor";
    private DetailedReportDao detailedReportDao;
    private ProjectAssignmentDao projectAssignmentDao;

    DetailedReportServiceImpl() {
    }

    @Autowired
    public DetailedReportServiceImpl(ReportCriteriaService reportCriteriaService, ProjectDao projectDao, TimesheetLockService lockService, DetailedReportDao detailedReportDao, ReportAggregatedDao reportAggregatedDAO, ProjectAssignmentDao projectAssignmentDao) {
        super(reportCriteriaService, projectDao, lockService, reportAggregatedDAO);
        this.detailedReportDao = detailedReportDao;
        this.projectAssignmentDao = projectAssignmentDao;
    }

    public ReportData getDetailedReportData(ReportCriteria reportCriteria) {
        return getReportData(reportCriteria);
    }

    @Override
    protected List<FlatReportElement> getReportElements(List<User> users,
                                                        List<Project> projects,
                                                        List<Date> lockedDates,
                                                        DateRange reportRange,
                                                        boolean showZeroBookings) {
        List<Integer> userIds = DomainUtil.getIdsFromDomainObjects(users);
        List<Integer> projectIds = DomainUtil.getIdsFromDomainObjects(projects);

        List<FlatReportElement> elements = getElements(userIds, projectIds, reportRange);

        for (FlatReportElement element : elements) {
            Date date = element.getDayDate();
            element.setLockableDate(new LockableDate(date, lockedDates.contains(date)));
        }

        if (showZeroBookings) {
            List<FlatReportElement> reportElementsForAssignmentsWithoutBookings = getReportElementsForAssignmentsWithoutBookings(reportRange, userIds, projectIds);

            reportElementsForAssignmentsWithoutBookings.addAll(elements);

            return reportElementsForAssignmentsWithoutBookings;
        } else {
            return elements;
        }
    }

    private List<FlatReportElement> getReportElementsForAssignmentsWithoutBookings(DateRange reportRange, List<Integer> userIds, List<Integer> projectIds) {
        List<ProjectAssignment> assignments = getAssignmentsWithoutBookings(reportRange, userIds, projectIds);

        List<FlatReportElement> elements = Lists.newArrayList();

        for (ProjectAssignment assignment : assignments) {
            elements.add(FlatReportElementBuilder.buildFlatReportElement(assignment));
        }

        return elements;
    }

    private List<FlatReportElement> getElements(List<Integer> userIds, List<Integer> projectIds, DateRange reportRange) {
        List<FlatReportElement> elements;

        if (userIds.isEmpty() && projectIds.isEmpty()) {
            elements = detailedReportDao.getHoursPerDay(reportRange);
        } else if (projectIds.isEmpty()) {
            elements = detailedReportDao.getHoursPerDayForUsers(userIds, reportRange);
        } else if (userIds.isEmpty()) {
            elements = detailedReportDao.getHoursPerDayForProjects(projectIds, reportRange);
        } else {
            elements = detailedReportDao.getHoursPerDayForProjectsAndUsers(projectIds, userIds, reportRange);
        }

        this.setCustomFields(elements);
        return elements;
    }

    private void setCustomFields(List<FlatReportElement> elements){
        if(elements != null && !elements.isEmpty()){
            Integer projectAssignmentId = null;
            for(FlatReportElement element:elements){
                if(!element.getCustomerCode().equals(CONTRACTOR_CUSTOMER_CODE)){
                    projectAssignmentId = element.getAssignmentId();
                    break;
                }
            }

            ProjectAssignment assignment = projectAssignmentDao.findById(projectAssignmentId);
            String sectionLeader = "";
            if(assignment.getProject().getSectionLeader() != null){
                sectionLeader = assignment.getProject().getSectionLeader().getFullName();
            }

            String headOfUnit = "";
            if(assignment.getProject().getHeadOfUnit() != null){
                headOfUnit = assignment.getProject().getHeadOfUnit().getFullName();
            }

            double allottedDays = 0.0;
            if(assignment.getAllottedHours() > 0){
                allottedDays = assignment.getAllottedHours() / 8;
            }

            for(FlatReportElement element:elements){
                element.setSectionLeader(sectionLeader);
                element.setHeadOfUnit(headOfUnit);
                element.setAssignmentDaysAllotted(allottedDays);
            }

        }
    }

}
