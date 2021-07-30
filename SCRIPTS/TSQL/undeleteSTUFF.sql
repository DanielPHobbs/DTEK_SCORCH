use Orchestrator

select * from POLICIES where Deleted = 1

use Orchestrator
UPDATE POLICIES Set Deleted = 0 where Deleted = 1