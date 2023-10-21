WITH vulnerability_cves AS (
    SELECT vulnerability_id, array_to_string(array_agg(reference), ', ') AS cves
    FROM dim_vulnerability_reference
    WHERE source = 'CVE'
    GROUP BY vulnerability_id
),
rolled_up_solutions AS (
    SELECT asset_id, vulnerability_id, COALESCE(dss.superceding_solution_id, davs.solution_id) AS solution_id
    FROM dim_asset_vulnerability_solution davs
    LEFT OUTER JOIN dim_solution_supercedence dss USING (solution_id)
),
cve_list AS (
    SELECT unnest(array[
        'CVE-2023-4663',
        'CVE-2023-331366',
        'CVE-2023-43323',
        'CVE-2022-41303'
        -- Add more CVEs here
    ]) AS cve
)

SELECT 
    ROUND(cvss_score::numeric, 1) AS "CVSS Score", 
    severity AS Severity, 
    da.ip_address AS "IP Address", 
    da.host_name AS "Host", 
    das.name AS Site, 
    dos.name AS "OS", 
    dos.version AS "OS Version", 
    dw.title AS "Vulnerability", 
    proofAsText(favi.proof) AS "Proof", 
    proofAsText(ds.fix) AS "Steps", 
    vcves.cves AS "CVE", 
    dv.title AS "Date Published", 
    da.last_assessed_for_vulnerabilities AS "Last Scanned"

FROM fact_asset_vulnerability_instance favi
JOIN dim_asset da USING (asset_id)
JOIN dim_site_asset dss USING (asset_id)
JOIN dim_operating_system dos USING (operating_system_id)
JOIN dim_vulnerability dw USING (vulnerability_id)
LEFT OUTER JOIN rolled_up_solutions sol USING (asset_id, vulnerability_id)
LEFT OUTER JOIN dim_solution ds USING (solution_id)
LEFT OUTER JOIN vulnerability_cves vcves USING (vulnerability_id)
WHERE 
    vcves.cves IN (SELECT cve FROM cve_list)
AND now() - da.last_assessed_for_vulnerabilities <= INTERVAL '15 days';
