SELECT DISTINCT
    om.OutreachStatus AS 'ORG Status',
    om.OutreachId AS 'OutreachID',
    om.OutreachTeam,
    pt.Name AS 'Project Type',
    ISNULL(SC.PhoneNum, '') AS 'Phone',
    CONVERT(VARCHAR(10), od.LastCallDate, 1) AS 'LastCall',
    oc.TotalCharts,
    oc.RetrievedCharts,
    oc.ToGoCharts,
    ISNULL(CONVERT(VARCHAR(10), om.ScheduleDate, 1), '') AS 'ScheduleDate',
    CONCAT(ISNULL(s.Address1, ''), ISNULL(s.Address2, '')) AS 'Address 1&2 Concat',
    ISNULL(s.City, '') AS 'City',
    ISNULL(s.Zip, '') AS 'Zip',
		(CASE WHEN s.[State] IN ('AK','ID','NV','HI','OR','WA','AZ','CA') THEN 'Pacific'
		WHEN s.[State] IN ('WY','MT','NM','UT','CO','TX') THEN 'Rockies'
		WHEN s.[State] IN ('CT','NJ') THEN 'Bridges & Tunnels'
		WHEN s.[State] = 'NY' THEN 'NY'
		WHEN s.[State] IN ('KS','AR','OK','MO') THEN 'Breadbasket'
		WHEN s.[State] IN ('SD','ND','IA','NE','MN','WI','IL') THEN 'Upper Midwest'
		WHEN s.[State] = 'FL' THEN 'FL'
		WHEN s.[State] IN ('GU','JA','AE','FT','PR','AA','VI','NA','MS','LA','AL') THEN 'The Bayou'
		WHEN s.[State] IS NULL THEN 'The Bayou'
		WHEN s.[State] IN ('MI','IN','OH') THEN 'The Rust Belt'
		WHEN s.[State] IN ('WV','KY','TN','PA') THEN 'Allegany'
		WHEN s.[State] IN ('SC','NC','GA') THEN 'Outer Banks'
		WHEN s.[State] IN ('DC','VT','ME','NH','RI','DE','MD','MA','VA') THEN 'Boston & The Beltway'
		ELSE 'The Bayou' END) AS 'Sub Region',
    ISNULL(pnp.Name, '') AS 'PNPCode',
    ISNULL(rm.Name, '') AS 'RetrievalMethod',
    CONVERT(VARCHAR(10), p.DueDate, 1) AS 'Project Due Date',
    ISNULL(om.HealthPortSiteID, '') AS 'ROISiteID',
    ISNULL(hr.SpecialHandlingCode, '') AS 'SH Code',
    ISNULL(hr.RDO, '') AS 'RDO',
    ISNULL(hr.VPO, '') AS 'VPO',
    ISNULL(hr.RMO, '') AS 'RMO',
    ISNULL(hr.SMP, '') AS 'SMP',
    ISNULL(p.Wave, '') AS 'Wave',
    ISNULL(s.SiteCleanId, '') AS 'SiteCleanId',
	DATEDIFF(day,om.InsertDate,GETDATE()) AS 'Days Since Creation'
 
FROM
    Chart c
    JOIN ProjectSite ps 
        ON c.ProjectId = ps.ProjectId 
        AND c.SiteId = ps.SiteId
    JOIN OutreachMaster om 
        ON ps.OutreachId = om.OutreachId
    LEFT JOIN OutreachCount oc 
        ON om.OutreachId = oc.OutreachID
    JOIN ProjectImport i 
        ON c.ChartId = i.ChartId
    LEFT JOIN OutreachDates od 
        ON om.OutreachId = od.OutreachId
    JOIN Project p 
        ON c.ProjectId = p.ProjectId
    LEFT JOIN [Site] s 
        ON om.PrimarySiteId = s.Id
    LEFT JOIN SiteContact sc 
        ON om.PrimarySiteId = sc.SiteID 
        AND sc.PrimaryFlag = '1' 
        AND ps.ProjectId = sc.ProjectId
	JOIN MasterNotes n
	   On om.OutreachID = n.OutreachID
    JOIN List pt 
        ON p.ProjectType = pt.Value 
        AND pt.ListType = 'ProjectType'
    LEFT JOIN List rm 
        ON om.RetrievalMethod = rm.Value 
        AND rm.ListType = 'RetrievalMethod'
    LEFT JOIN List pnp 
        ON om.PNPCode = pnp.Value 
        AND pnp.ListType = 'PNPCode'
    LEFT JOIN List a 
        ON p.AuditType = a.Value 
        AND a.ListType = 'AuditType'
    LEFT JOIN List emr 
        ON om.EMRSystem = emr.Value 
        AND emr.ListType = 'EMRSystem'
    LEFT JOIN HealthPort_RDO hr 
        ON om.HealthPortSiteId = CAST(hr.SiteId AS VARCHAR(50))
 
WHERE
    DATEDIFF(day, GETDATE(), p.DueDate) > 1 
    AND om.HealthPortSiteId IS NULL 
    AND hr.RDO IS NULL 
    AND hr.SpecialHandlingCode IS NULL 
    AND om.OutreachStatus IN ('Unscheduled', 'PNP Released', 'Escalated') 
    AND rm.Name NOT IN ('CIOX', 'CIOX - Onsite', 'CIOX - Partner', 'CIOX - Remote') 
    AND rm.Name NOT LIKE 'HIH%' 
    AND a.Name IN ('Medicare Risk', 'Medicaid Risk', 'RADV', 'HEDIS', 'ACA') 
    AND pt.Name NOT LIKE '%Digital Direct'
	AND n.NoteType Not IN ('114','115','116') 
	AND (pnp.Name NOT IN ('PNP007', 'PNP011', 'PNP024', 'PNP027', 'PNP026', 'PNP028', 'PNP039', 'PNP042', 'PNP051') OR pnp.Name IS NULL)

 
Group By
    om.OutreachStatus,
    om.OutreachId,
    om.OutreachTeam,
    pt.Name,
    ISNULL(SC.PhoneNum, ''),
    CONVERT(VARCHAR(10), od.LastCallDate, 1),
    oc.TotalCharts,
    oc.RetrievedCharts,
    oc.ToGoCharts,
    ISNULL(CONVERT(VARCHAR(10), om.ScheduleDate, 1), ''),
    CONCAT(ISNULL(s.Address1, ''), ISNULL(s.Address2, '')),
    ISNULL(s.City, ''),
    ISNULL(s.Zip, ''),
	(CASE WHEN s.[State] IN ('AK','ID','NV','HI','OR','WA','AZ','CA') THEN 'Pacific'
		WHEN s.[State] IN ('WY','MT','NM','UT','CO','TX') THEN 'Rockies'
		WHEN s.[State] IN ('CT','NJ') THEN 'Bridges & Tunnels'
		WHEN s.[State] = 'NY' THEN 'NY'
		WHEN s.[State] IN ('KS','AR','OK','MO') THEN 'Breadbasket'
		WHEN s.[State] IN ('SD','ND','IA','NE','MN','WI','IL') THEN 'Upper Midwest'
		WHEN s.[State] = 'FL' THEN 'FL'
		WHEN s.[State] IN ('GU','JA','AE','FT','PR','AA','VI','NA','MS','LA','AL') THEN 'The Bayou'
		WHEN s.[State] IS NULL THEN 'The Bayou'
		WHEN s.[State] IN ('MI','IN','OH') THEN 'The Rust Belt'
		WHEN s.[State] IN ('WV','KY','TN','PA') THEN 'Allegany'
		WHEN s.[State] IN ('SC','NC','GA') THEN 'Outer Banks'
		WHEN s.[State] IN ('DC','VT','ME','NH','RI','DE','MD','MA','VA') THEN 'Boston & The Beltway'
		ELSE 'The Bayou' END),
    ISNULL(pnp.Name, ''),
    ISNULL(rm.Name, ''),
    CONVERT(VARCHAR(10), p.DueDate, 1),
    ISNULL(om.HealthPortSiteID, ''),
    ISNULL(hr.SpecialHandlingCode, ''),
    ISNULL(hr.RDO, ''),
    ISNULL(hr.VPO, ''),
    ISNULL(hr.RMO, ''),
    ISNULL(hr.SMP, ''),
    ISNULL(p.Wave, ''),
    ISNULL(s.SiteCleanId, ''),
	om.InsertDate
 
ORDER BY
    pt.Name ASC,
    CONCAT(ISNULL(s.Address1, ''), ISNULL(s.Address2, '')) ASC;