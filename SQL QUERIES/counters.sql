Select Ob.Name, C.DefaultValue from COUNTERS C
inner join OBJECTS Ob on Ob.UniqueID = C.UniqueID