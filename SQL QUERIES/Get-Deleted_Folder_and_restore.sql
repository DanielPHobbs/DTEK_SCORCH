Select UniqueID, ParentID, LastModified, Deleted from FOLDERS where Name like 'My deleted folder name'

UPDATE FOLDERS set Deleted = 0 where UniqueID = 'The UniqueID of the deleted folder, you can get it by running the "Find a deleted folder SQL query"'