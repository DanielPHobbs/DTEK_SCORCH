use [SCORCHPersistantDB]

/*Data for the table StateTracking02 */
  
INSERT INTO StateTracking02(
guid,
runcount,
runbookname,
description,
Activityname,
Activitystatus,
activitystart,
activityend,
errorseverity,
errorreason,
Alltaskstatus,
runbookserver,
displayname) 

VALUES (
'd3995e1f-a9f7-43e6-8922-9d843e72ceb9',
'1',
'CreateAccount',
'Create user account in active directory',
'create user',
'success',
'12/10/2021 21:02:48',
'12/10/2021 21:04:48',
'No Error',
'No reason',
'success',
'dtekorch16-s2',
'alec guiness'
);
-- repeat again for additional row
