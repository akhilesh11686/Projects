SELECT 
       DPH.JBHDR_JOB_REF AS BKG,
       DPH.ADDPT_DEPT_CODE_DEFAULT,
       dc.ADDPT_DEPT_CODE,
       DECODE(DC.DDCHG_CHARGE_STATUS, '', 'No Charges', 'NOC', 'No Charges', 'REJ', 'Rejected', 'INV', 'Original', 'TBI', 'Full Charges', 'DRA' , 'Proforma', 'ONH' , 'On Hold', 'Z' , 'Calculating') AS CHG_STATUS,
       DECODE(DPH.DDTRG_IMPORT_EXPORT_FLAG, 'E', 'Export', 'I', 'Import') AS IMPORT_EXPORT_FLAG,
       DECODE(DC.DDCHG_CHARGE_TYPE, '', 'No Tariff Found','T', 'Detention', 'S', 'Storage', 'M', 'Demurrage', 'G','Merged','R', 'Monitoring') AS DD_CHARGE_TYPE,
       DPH.DDPRD_EQUIPMENT_NUMBER AS CONTAINER,
       RM.RESOURCE_SIZE AS CONT_SIZE,
       RM.RESOURCE_TYPE AS CONT_TYPE,
       DPH.VOVOY_VOYAGE_REF AS VOYAGE, 
       DPH.DDPRD_START_STATUS AS START_STAT,
       DPH.DDPRD_START_LOCATION AS START_LOC,
       TO_CHAR(DPH.DDPRD_START_DATE_TIME, 'MM/DD/YYYY') AS START_DATE,
       TO_CHAR(DPH.DDPRD_START_DATE_TIME, 'HH24:MI:SS') AS START_TIME,
       DC.DDCHG_START_DATE AS START_CHG_DATE,
       DECODE(DC.DDCHG_START_OVERRIDE_TYPE, 'CS', 'Voyage Start Day', 'SD', 'Container Start Day') AS START_OVERRIDE_RULE,
       DECODE(DPH.DDPRD_START_OVERRIDE_FLAG, 'N', 'No', 'Y', 'Yes') AS START_MOVE_UPDATE,
       DPH.DDPRD_STOP_STATUS AS STOP_STAT,
       DPH.DDPRD_STOP_LOCATION AS STOP_LOC,
       TO_CHAR(DPH.DDPRD_STOP_DATE_TIME, 'MM/DD/YYYY') AS STOP_DATE,
       TO_CHAR(DPH.DDPRD_STOP_DATE_TIME, 'HH24:MI:SS') AS STOP_TIME,
       DC.DDCHG_STOP_DATE AS STOP_CHG_DATE,
       DECODE(DC.DDCHG_STOP_OVERRIDE_TYPE, 'CS', 'Voyage Cut Day', 'SD', 'Container Stop Day') AS STOP_OVERRIDE_RULE,
       DECODE(DPH.DDPRD_STOP_OVERRIDE_FLAG, 'N', 'No', 'Y', 'Yes') AS STOP_MOVE_UPDATE,
       DC.DDCHG_REASON_CODE,
       DC.DDCHG_PAYER_CODE,
       jrr.created_date "Rollover Date",
       jrr.created_by "Rollover Created By",
       jrr.jbrre_reason_code "Rollover Reason Code",
       c.description "Rollover Reason",
       jrr.vovoy_voyage_reference_old "Old Voyage Reference",
       jrr.vovoy_voyage_reference_new "New Voyage Reference",
       jrr.jbrre_comments "Rollover Comments",
       c.next_value "Fault",
       vp1.cutoff_date "Old Committed Date",
       vp2.cutoff_date "New Committed Date"       
FROM 
	   DDPRD_PERIOD DPH  
       INNER JOIN RM_EQUIPMENT RM ON RM.RM_RESOURCE_CODE = DPH.DDPRD_EQUIPMENT_NUMBER                
       JOIN DDPAC_PERIOD_CHARGE DPC ON DPC.DDPRD_PERIOD_ID = DPH.DDPRD_PERIOD_ID
       inner JOIN DDCHG_CHARGE DC ON DC.DDCHG_CHARGE_ID = DPC.DDCHG_CHARGE_ID
       inner JOIN PARTNERS P ON P.PARTNER_CODE = DC.DDCHG_PAYER_CODE
       inner JOIN JBRRE_ROLLOVER_REASONS JRR ON JRR.JBHDR_JOB_REFERENCE = dph.jbhdr_job_ref 
       inner JOIN codes c ON c.code_value = jrr.jbrre_reason_code AND c.code_type = 'ROLL'
       LEFT OUTER JOIN voyage_ports vp1 ON vp1.voyage_reference = jrr.vovoy_voyage_reference_old AND vp1.pool_location = jrr.jbrre_old_voy_pool_loc AND vp1.schedule_type = 'OPS' AND vp1.omit_flag <> 'Y'
       LEFT OUTER JOIN voyage_ports vp2 ON vp2.voyage_reference = jrr.vovoy_voyage_reference_new AND vp2.pool_location = jrr.jbrre_new_voy_pool_loc AND vp2.schedule_type = 'OPS' AND vp2.omit_flag <> 'Y'
WHERE DC.DDCHG_CHARGE_TYPE = 'S' AND DPH.DDPRD_START_STATUS ='XRX' AND DC.DDCHG_CHARGE_STATUS = 'TBI' AND substr(DPH.DDPRD_START_LOCATION,1,2) in ('US','CA') AND DPH.SHIPCOMP_CODE IN ('0001','0002')
	AND DPH.DDPRD_START_DATE_TIME between trunc(sysdate) - 30 and trunc(sysdate) - 14 
ORDER BY dph.jbhdr_job_ref, jrr.created_date
	
