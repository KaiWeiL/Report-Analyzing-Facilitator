SELECT EVENT_DT AS "Event Date/Time", AGENTNAME AS "Agent Name", HOSTNAME AS "Hostname",
AGENTTYPE_D AS "Agent Type", HOSTADDR AS "IP Address", EVENT_TYPE_D AS "Event Type",
Description AS "Description", EVENT_SEVERITY_D AS "Severity",
EVENT_PRIORITY AS "Event Priority", DISPOSITION_D AS "Disposition",
AGENT_VERSION AS "Agent Version", OSTYPE_D AS "OS Version", EVENT_CNT AS "Event Count",
EVENT_DURATION AS "Event Duration", POST_DT AS "Post Date/Time",
RULE_NAME AS "Rule Name", USER_NAME AS "User Name", DOMAIN_NAME AS "Domain Name",
SYSTEM_STATE_D AS "Policy Overridden", OPERATION_D AS "Operation",
TARGET_INFO AS "Resource", PROCESS_PATH AS "Process",
PROCESS_EFA_PUB AS "Process Publisher", PROCESS_EFA_FLAG AS "Process Signature"
FROM CSPEVENT_VW
WITH (NOLOCK)
ORDER BY EVENT_DT DESC, EVENT_SEQ DESC
