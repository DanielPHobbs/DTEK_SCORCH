CREATE TABLE StateTracking01 (
    Guid			varchar(255) NOT NULL,
	RunCount		int,
    RunbookName		varchar(255),
	Description		varchar(255),
	AccountName		varchar(255),
	HomedriveMap	varchar(255),
	ExchangeMB		varchar(255),
	GroupMembership varchar(255),
    ActivityName	varchar(255),
	ActivityStatus	DATETIME,
	AvctivityStart	DATETIME,
	ActivityEnd		DATETIME,
	ErrorSeverity	varchar(255),
	ErrorReason		varchar(255),
	AllTaskStatus	varchar(255),
	RunbookServer	varchar(255),
	ExtendedData1	varchar(255),
    ExtendedData2	varchar(255),
    ExtendedData3	varchar(255),
	PRIMARY KEY (Guid)
);

/*
DROP TABLE StateTracking01

ALTER TABLE StateTracking01
  ADD last_name VARCHAR(50),
	  first_name VARCHAR(50)

ALTER TABLE StateTracking01
  ALTER COLUMN last_name VARCHAR(75) NOT NULL;

ALTER TABLE StateTracking01
  DROP COLUMN last_name;
 */