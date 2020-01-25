package net.rrm.ehour.report.reports.element;

import net.rrm.ehour.domain.Customer;
import net.rrm.ehour.domain.Project;
import net.rrm.ehour.domain.ProjectAssignment;
import net.rrm.ehour.domain.User;

public class FlatReportElementBuilder {
    private FlatReportElementBuilder() {
    }

    public static FlatReportElement buildFlatReportElement(ProjectAssignment assignment) {
        FlatReportElement element = new FlatReportElement();

        element.setAssignmentId(assignment.getAssignmentId());
        element.setRole(assignment.getRole());
        element.setInternalAddress(assignment.getInternalAddress());
        element.setTelephone(assignment.getTelephone());
        if(assignment.getAllottedHours() != null && assignment.getAllottedHours() > 0){
            double daysAlloted = assignment.getAllottedHours() / 8;
            element.setAssignmentDaysAllotted(daysAlloted);
            element.setAcumulatedAssignmentDaysAllotted(daysAlloted);
        }
        element.setAssignmentEndDate(assignment.getDateEnd());

        Project project = assignment.getProject();
        Customer customer = project.getCustomer();

        element.setCustomerCode(customer.getCode());
        element.setCustomerId(customer.getCustomerId());
        element.setCustomerName(customer.getName());

        element.setEmptyEntry(true);
        element.setProjectCode(project.getProjectCode());
        element.setProjectId(project.getProjectId());
        element.setProjectName(project.getName());
        if(project.getSectionLeader() != null){
            element.setSectionLeader(project.getSectionLeader().getFullName());
        }
        if(project.getHeadOfUnit() != null){
            element.setHeadOfUnit(project.getHeadOfUnit().getFullName());
        }
        if(project.getContractManager() != null){
            element.setContractManager(project.getContractManager().getFullName());
        }

        element.setRate(assignment.getHourlyRate());

        User user = assignment.getUser();
        element.setUserId(user.getUserId());
        element.setUserFirstName(user.getFirstName());
        element.setUserLastName(user.getLastName());

        element.setDisplayOrder(1);

        return element;
    }
}
