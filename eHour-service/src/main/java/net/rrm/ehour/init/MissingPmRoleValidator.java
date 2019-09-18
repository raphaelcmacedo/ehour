package net.rrm.ehour.init;

import net.rrm.ehour.domain.Project;
import net.rrm.ehour.domain.User;
import net.rrm.ehour.domain.UserRole;
import net.rrm.ehour.persistence.project.dao.ProjectDao;
import net.rrm.ehour.persistence.user.dao.UserDao;
import net.rrm.ehour.project.service.ProjectService;
import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.PostConstruct;
import java.util.List;

/**
 * EHO-381 The PM role was lost when a user with an existing PM role was edited.
 * This validator re-adds the PM role to users which are marked as a PM but don't have this role.
 */
@Service
public class MissingPmRoleValidator {
    private static final Logger LOGGER = Logger.getLogger(MissingPmRoleValidator.class);

    @Autowired
    private ProjectDao projectDao;

    @Autowired
    private ProjectService projectService;

    @Autowired
    private UserDao userDao;

    @PostConstruct
    @Transactional(propagation = Propagation.REQUIRES_NEW)
    public void init() {
        LOGGER.info("Finding and fixing users who are PM but don't have PM role (EHO-381)");

        List<Project> projectsWithPmSet = projectDao.findAllProjectsWithPmSet();

        for (Project project : projectsWithPmSet) {
            this.applyPMRole(project.getProjectManager(), project);
            this.applyPMRole(project.getSectionLeader(), project);
            this.applyPMRole(project.getHeadOfUnit(), project);
            this.applyPMRole(project.getContractManager(), project);
        }
    }

    @Transactional
    private void applyPMRole(User user, Project project){
        if (user != null) {
            boolean hasPmRole = user.getUserRoles().contains(UserRole.PROJECTMANAGER);

            if (!hasPmRole) {
                LOGGER.warn(String.format("%s (%s) does not have PM role but is PM for %s. Adding PM role.", user.getFullName(), user.getPK(), project.getFullName()));
                projectService.validatePMRoles(project);
            }
        }
    }
}
