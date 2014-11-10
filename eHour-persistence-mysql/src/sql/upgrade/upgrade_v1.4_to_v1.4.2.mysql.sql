CREATE TABLE USER_TO_DEPARTMENT (
  DEPARTMENT_ID INT(11) NOT NULL,
  USER_ID       INT(11) NOT NULL,
  PRIMARY KEY (DEPARTMENT_ID, USER_ID),
  KEY ROLE (DEPARTMENT_ID),
  KEY USER_ID (USER_ID),
  CONSTRAINT FK_USER_TO_USER FOREIGN KEY (USER_ID) REFERENCES USERS (USER_ID),
  CONSTRAINT FK_USER_TO_DEPT FOREIGN KEY (DEPARTMENT_ID) REFERENCES USER_DEPARTMENT (DEPARTMENT_ID)
)
  ENGINE =InnoDB
  DEFAULT CHARSET =utf8;


ALTER TABLE USER_DEPARTMENT ADD COLUMN MANAGER_USER_ID INT(11) DEFAULT NULL;
ALTER TABLE USER_DEPARTMENT ADD COLUMN TIMEZONE VARCHAR(128) DEFAULT NULL;
ALTER TABLE USER_DEPARTMENT ADD COLUMN PARENT_DEPARTMENT_ID INT(11) DEFAULT NULL;

ALTER TABLE USERS MODIFY COLUMN DEPARTMENT_ID INT NULL;

UPDATE CONFIGURATION SET CONFIG_VALUE = '1.4.2' WHERE CONFIG_KEY = 'version';

