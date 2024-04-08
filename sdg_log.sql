create table sdg_log (
	sdg_id			number(16)		not null,
	time			date			not null,
	application_code	varchar2(255)		not null,
	session_id		number(16),
	description		varchar2(4000),
	constraint PK_sdg_log primary key (sdg_id, time, application_code)
	      using index
	      tablespace LIMS_INDEX
)
tablespace LIMS_TABLE
/

create index ix_sdg_log_code on sdg_log (
   application_code ASC,
   time asc
)
tablespace LIMS_INDEX
/

GRANT SELECT , INSERT , UPDATE , DELETE ON sdg_log TO LIMS_SYS WITH GRANT OPTION
/

CREATE OR REPLACE FORCE VIEW LIMS_SYS.sdg_log
	(sdg_id, time, application_code, session_id, description)
AS 
SELECT sdg_id, time, application_code, session_id, description
FROM LIMS.sdg_log
/

GRANT SELECT ON  LIMS_SYS.sdg_log TO LIMS_READONLY
/

GRANT DELETE, INSERT, SELECT, UPDATE ON  LIMS_SYS.sdg_log TO LIMS_USER
/
