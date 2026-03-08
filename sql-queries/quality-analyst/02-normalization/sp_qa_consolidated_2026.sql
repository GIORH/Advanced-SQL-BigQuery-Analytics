/*
===============================================================================
PROJECT: Enterprise Global Quality PPO Engine
LAYER: Silver / Normalization
ROLE: Quality Analyst (Entry Level)
DESCRIPTION: 
    This Stored Procedure consolidates 7 distinct data sources into a 
    standardized analytical table. It handles data cleaning, metric 
    weighting from a master table, and multi-stage performance scoring.

DATA SOURCES:
    1. VOC (Voice of the Customer)
    2. Internal Quality Audits
    3. Team Development (Workshops)
    4. PPD (Professional Development)
    5. Productivity (Evaluations & Calibrations)
    6. Accuracy (Dispute & Audit Precision)
    7. Targets Master (Dynamic Weights/Goals)

AUTHOR: GIORH
DATE: 2026-03-08
===============================================================================
*/

CREATE OR REPLACE PROCEDURE `ti-ca-ml-start.QA_Training_scorecards.sp_2026_scorecard_quality_analyst_consolidated`()
BEGIN

  -- Create or Refresh the Consolidated Silver Table
  CREATE OR REPLACE TABLE `ti-ca-ml-start.QA_Training_scorecards.2026_scorecard_quality_analyst_consolidated` AS
  WITH 
  
  -- STEP 1: Fetch dynamic metric weights based on position and current date
  weights AS (
    SELECT metric_name, metric_weight
    FROM `ti-ca-ml-start.QA_Training_scorecards.targets_master`
    WHERE position = 'CX Quality Analyst'
      AND CURRENT_DATE() BETWEEN SAFE_CAST(effective_from AS DATE) AND SAFE_CAST(effective_to AS DATE)
  ),

  -- STEP 2: RAW DATA CLEANING & TYPE CASTING
  -- Standardizing all 6 primary sources to ensure consistent JOIN keys
  voc_raw AS (
    SELECT DISTINCT
      report_month, SAFE_CAST(report_year AS INT64) as report_year, SAFE_CAST(tm_wdid AS INT64) as tm_wdid,
      tm_name, program, country, region, SAFE_CAST(tl_wdid AS INT64) as tl_wdid,
      tl_name, tl_manager, tl_sr_manager,
      SAFE_CAST(REPLACE(voc_score_actual, '%', '') AS FLOAT64) as voc_actual,
      SAFE_CAST(REPLACE(voc_score_target_program, '%', '') AS FLOAT64) as voc_target_program,
      notes as voc_notes, kpi_applicable as voc_applicable, exempt as voc_exempt, reason as voc_reason
    FROM `ti-ca-ml-start.QA_Training_scorecards.2026_qa_voc_raw`
    WHERE tm_wdid IS NOT NULL AND tm_wdid NOT IN ('tm_wdid', 'tm_name')
  ),

  quality_raw AS (
    SELECT DISTINCT
      report_month, SAFE_CAST(report_year AS INT64) as report_year, SAFE_CAST(tm_wdid AS INT64) as tm_wdid,
      tm_name, program, country, region, SAFE_CAST(tl_wdid AS INT64) as tl_wdid,
      tl_name, tl_manager, tl_sr_manager,
      SAFE_CAST(REPLACE(quality_score_actual, '%', '') AS FLOAT64) as qa_actual,
      SAFE_CAST(REPLACE(quality_score_target_program, '%', '') AS FLOAT64) as qa_target_program,
      notes as qa_notes, kpi_applicable as qa_applicable, exempt as qa_exempt, reason as qa_reason
    FROM `ti-ca-ml-start.QA_Training_scorecards.2026_qa_quality_raw`
    WHERE tm_wdid IS NOT NULL AND tm_wdid NOT IN ('tm_wdid', 'tm_name')
  ),

  team_dev_raw AS (
    SELECT DISTINCT
      report_month, SAFE_CAST(report_year AS INT64) as report_year, SAFE_CAST(tm_wdid AS INT64) as tm_wdid,
      tm_name, program, country, region, SAFE_CAST(tl_wdid AS INT64) as tl_wdid,
      tl_name, tl_manager, tl_sr_manager,
      SAFE_DIVIDE(SAFE_CAST(workshops_delivered AS FLOAT64), SAFE_CAST(workshops_required AS FLOAT64)) * 100 as td_actual,
      notes as td_notes, kpi_applicable as td_applicable, exempt as td_exempt, reason as td_reason
    FROM `ti-ca-ml-start.QA_Training_scorecards.2026_qa_team_dev_raw`
    WHERE tm_wdid IS NOT NULL AND tm_wdid NOT IN ('tm_wdid', 'tm_name')
  ),

  ppd_raw AS (
    SELECT DISTINCT
      report_month, SAFE_CAST(report_year AS INT64) as report_year, SAFE_CAST(tm_wdid AS INT64) as tm_wdid,
      tm_name, program, country, region, SAFE_CAST(tl_wdid AS INT64) as tl_wdid,
      tl_name, tl_manager, tl_sr_manager,
      SAFE_DIVIDE(SAFE_CAST(courses_completed AS FLOAT64), SAFE_CAST(courses_required AS FLOAT64)) * 100 as ppd_actual,
      notes as ppd_notes, kpi_applicable as ppd_applicable, exempt as ppd_exempt, reason as ppd_reason
    FROM `ti-ca-ml-start.QA_Training_scorecards.2026_qa_ppd_raw`
    WHERE tm_wdid IS NOT NULL AND tm_wdid NOT IN ('tm_wdid', 'tm_name')
  ),

  productivity_raw AS (
    SELECT DISTINCT
      report_month, SAFE_CAST(report_year AS INT64) as report_year, SAFE_CAST(tm_wdid AS INT64) as tm_wdid,
      tm_name, program, country, region, SAFE_CAST(tl_wdid AS INT64) as tl_wdid,
      tl_name, tl_manager, tl_sr_manager,
      SAFE_CAST(qa_evals_completed AS FLOAT64) as qa_evals_completed,
      SAFE_CAST(qa_evals_target AS FLOAT64) as qa_evals_target,
      SAFE_CAST(voc_evals_completed AS FLOAT64) as voc_evals_completed,
      SAFE_CAST(voc_evals_target AS FLOAT64) as voc_evals_target,
      SAFE_CAST(calibrations_completed AS FLOAT64) as calibrations_completed,
      SAFE_CAST(calibrations_target AS FLOAT64) as calibrations_target,
      qa_evals_na, voc_na, calibrations_na,
      notes as prod_notes, kpi_applicable as prod_applicable, exempt as prod_exempt, reason as prod_reason
    FROM `ti-ca-ml-start.QA_Training_scorecards.2026_qa_productivity_raw`
    WHERE tm_wdid IS NOT NULL AND tm_wdid NOT IN ('tm_wdid', 'tm_name')
  ),

  accuracy_raw AS (
    SELECT DISTINCT
      report_month, SAFE_CAST(report_year AS INT64) as report_year, SAFE_CAST(tm_wdid AS INT64) as tm_wdid,
      tm_name, program, country, region, SAFE_CAST(tl_wdid AS INT64) as tl_wdid,
      tl_name, tl_manager, tl_sr_manager,
      SAFE_CAST(cal_accurate AS FLOAT64) as cal_accurate,
      SAFE_CAST(cal_performed AS FLOAT64) as cal_performed,
      SAFE_CAST(valid_disputes AS FLOAT64) as valid_disputes,
      SAFE_CAST(total_completed_evals AS FLOAT64) as total_completed_evals,
      SAFE_CAST(accurate_evals AS FLOAT64) as accurate_evals,
      notes as acc_notes, kpi_applicable as acc_applicable, exempt as acc_exempt, reason as acc_reason,
      calibrations_na as acc_calibrations_na, disputes_na as acc_disputes_na, audits_na as acc_audits_na
    FROM `ti-ca-ml-start.QA_Training_scorecards.2026_qa_accuracy_raw`
    WHERE tm_wdid IS NOT NULL AND tm_wdid NOT IN ('tm_wdid', 'tm_name')
  ),

  -- STEP 3: BASE CONSOLIDATION
  -- Using FULL OUTER JOIN to prevent data loss across disparate sources
  base AS (
    SELECT 
      COALESCE(v.tm_wdid, q.tm_wdid, t.tm_wdid, p.tm_wdid, pr.tm_wdid, a.tm_wdid) as employee_wdid,
      COALESCE(v.tm_name, q.tm_name, t.tm_name, p.tm_name, pr.tm_name, a.tm_name) as employee_name,
      COALESCE(v.report_month, q.report_month, t.report_month, p.report_month, pr.report_month, a.report_month) as report_month,
      COALESCE(v.report_year, q.report_year, t.report_year, p.report_year, pr.report_year, a.report_year) as report_year,
      COALESCE(v.program, q.program, t.program, p.program, pr.program, a.program) as program,
      COALESCE(v.country, q.country, t.country, p.country, pr.country, a.country) as country,
      COALESCE(v.region, q.region, t.region, p.region, pr.region, a.region) as region,
      COALESCE(v.tl_wdid, q.tl_wdid, t.tl_wdid, p.tl_wdid, pr.tl_wdid, a.tl_wdid) as tl_wdid,
      COALESCE(v.tl_name, q.tl_name, t.tl_name, p.tl_name, pr.tl_name, a.tl_name) as tl_name,
      TRIM(REGEXP_EXTRACT(COALESCE(v.tl_manager, q.tl_manager, t.tl_manager, p.tl_manager, pr.tl_manager, a.tl_manager), r'^[^\(]+')) as tl_manager_name,
      SAFE_CAST(REGEXP_EXTRACT(COALESCE(v.tl_manager, q.tl_manager, t.tl_manager, p.tl_manager, pr.tl_manager, a.tl_manager), r'\((\d+)\)') AS INT64) as tl_manager_wdid,
      TRIM(REGEXP_EXTRACT(COALESCE(v.tl_sr_manager, q.tl_sr_manager, t.tl_sr_manager, p.tl_sr_manager, pr.tl_sr_manager, a.tl_sr_manager), r'^[^\(]+')) as tl_sr_manager_name,
      SAFE_CAST(REGEXP_EXTRACT(COALESCE(v.tl_sr_manager, q.tl_sr_manager, t.tl_sr_manager, p.tl_sr_manager, pr.tl_sr_manager, a.tl_sr_manager), r'\((\d+)\)') AS INT64) as tl_sr_manager_wdid,
      
      q.qa_actual, q.qa_target_program, q.qa_exempt, q.qa_notes, q.qa_applicable, q.qa_reason,
      v.voc_actual, v.voc_target_program, v.voc_exempt, v.voc_notes, v.voc_applicable, v.voc_reason,
      t.td_actual, t.td_exempt, t.td_notes, t.td_applicable, t.td_reason,
      p.ppd_actual, p.ppd_exempt, p.ppd_notes, p.ppd_applicable, p.ppd_reason,
      pr.qa_evals_completed, pr.qa_evals_target, pr.voc_evals_completed, pr.voc_evals_target, pr.calibrations_completed, pr.calibrations_target, pr.qa_evals_na, pr.voc_na, pr.calibrations_na, pr.prod_notes, pr.prod_applicable, pr.prod_exempt, pr.prod_reason,
      a.cal_accurate, a.cal_performed, a.valid_disputes, a.total_completed_evals, a.accurate_evals, a.acc_notes, a.acc_applicable, a.acc_exempt, a.acc_reason, a.acc_calibrations_na, a.acc_disputes_na, a.acc_audits_na,

      (SELECT metric_weight FROM weights WHERE metric_name = 'Quality % to Goal') as qa_weight,
      (SELECT metric_weight FROM weights WHERE metric_name = 'VOC % to Goal') as voc_weight,
      (SELECT metric_weight FROM weights WHERE metric_name = 'Team Development') as td_weight,
      (SELECT metric_weight FROM weights WHERE metric_name = 'PPD') as ppd_weight,
      (SELECT metric_weight FROM weights WHERE metric_name = 'Productivity') as prod_weight,
      (SELECT metric_weight FROM weights WHERE metric_name = 'Accuracy % to Goal') as accuracy_weight
    FROM voc_raw v
    FULL OUTER JOIN quality_raw q ON v.tm_wdid = q.tm_wdid AND v.report_month = q.report_month AND v.program = q.program
    FULL OUTER JOIN team_dev_raw t ON COALESCE(v.tm_wdid, q.tm_wdid) = t.tm_wdid AND COALESCE(v.report_month, q.report_month) = t.report_month AND COALESCE(v.program, q.program) = t.program
    FULL OUTER JOIN ppd_raw p ON COALESCE(v.tm_wdid, q.tm_wdid, t.tm_wdid) = p.tm_wdid AND COALESCE(v.report_month, q.report_month, t.report_month) = p.report_month AND COALESCE(v.program, q.program, t.program) = p.program
    FULL OUTER JOIN productivity_raw pr ON COALESCE(v.tm_wdid, q.tm_wdid, t.tm_wdid, p.tm_wdid) = pr.tm_wdid AND COALESCE(v.report_month, q.report_month, t.report_month, p.report_month) = pr.report_month AND COALESCE(v.program, q.program, t.program, p.program) = pr.program
    FULL OUTER JOIN accuracy_raw a ON COALESCE(v.tm_wdid, q.tm_wdid, t.tm_wdid, p.tm_wdid, pr.tm_wdid) = a.tm_wdid AND COALESCE(v.report_month, q.report_month, t.report_month, p.report_month, pr.report_month) = a.report_month AND COALESCE(v.program, q.program, t.program, p.program, pr.program) = a.program
  ),

  -- STEP 4: BUSINESS LOGIC - COMPONENT COUNTS
  -- Determining the denominator for multi-component metrics (Productivity/Accuracy)
  stage_1 AS (
    SELECT *,
      CASE WHEN UPPER(TRIM(prod_applicable)) = 'YES' AND UPPER(TRIM(prod_exempt)) = 'NO' THEN (COALESCE((CASE WHEN UPPER(TRIM(qa_evals_na)) = 'YES' THEN 1 ELSE 0 END), 0) + COALESCE((CASE WHEN UPPER(TRIM(voc_na)) = 'YES' THEN 1 ELSE 0 END), 0) + COALESCE((CASE WHEN UPPER(TRIM(calibrations_na)) = 'YES' THEN 1 ELSE 0 END), 0)) ELSE NULL END AS productivity_applicable_components_count,
      CASE WHEN UPPER(TRIM(acc_applicable)) = 'YES' AND UPPER(TRIM(acc_exempt)) = 'NO' THEN (COALESCE((CASE WHEN UPPER(TRIM(acc_calibrations_na)) = 'YES' THEN 1 ELSE 0 END), 0) + COALESCE((CASE WHEN UPPER(TRIM(acc_disputes_na)) = 'YES' THEN 1 ELSE 0 END), 0) + COALESCE((CASE WHEN UPPER(TRIM(acc_audits_na)) = 'YES' THEN 1 ELSE 0 END), 0)) ELSE NULL END AS accuracy_applicable_components_count
    FROM base
  ),

  -- STEP 5: PERFORMANCE CALCULATION - INDIVIDUAL RATES
  stage_2 AS (
    SELECT *,
      CASE WHEN productivity_applicable_components_count >= 0 AND UPPER(TRIM(qa_evals_na)) = 'YES' THEN SAFE_DIVIDE(qa_evals_completed, qa_evals_target) * 100 ELSE NULL END AS qa_evals_completion,
      CASE WHEN productivity_applicable_components_count >= 0 AND UPPER(TRIM(voc_na)) = 'YES' THEN SAFE_DIVIDE(voc_evals_completed, voc_evals_target) * 100 ELSE NULL END AS voc_evals_completion,
      CASE WHEN productivity_applicable_components_count >= 0 AND UPPER(TRIM(calibrations_na)) = 'YES' THEN SAFE_DIVIDE(calibrations_completed, calibrations_target) * 100 ELSE NULL END AS calibrations_completion,
      CASE WHEN accuracy_applicable_components_count >= 0 AND UPPER(TRIM(acc_calibrations_na)) = 'YES' THEN SAFE_DIVIDE(cal_accurate, cal_performed) * 100 ELSE NULL END AS cal_accuracy_rate,
      CASE WHEN accuracy_applicable_components_count >= 0 AND UPPER(TRIM(acc_disputes_na)) = 'YES' THEN SAFE_DIVIDE(valid_disputes, total_completed_evals) * 100 ELSE NULL END AS dispute_rate,
      CASE WHEN accuracy_applicable_components_count >= 0 AND UPPER(TRIM(acc_audits_na)) = 'YES' THEN SAFE_DIVIDE(accurate_evals, total_completed_evals) * 100 ELSE NULL END AS eval_accuracy_rate
    FROM stage_1
  ),

  -- STEP 6: PERCENT TO GOAL (PTG) MAPPING
  stage_3 AS (
    SELECT *,
      SAFE_DIVIDE(COALESCE(qa_evals_completion, 0) + COALESCE(voc_evals_completion, 0) + COALESCE(calibrations_completion, 0), productivity_applicable_components_count) as productivity_pct_to_goal,
      CASE WHEN cal_accuracy_rate IS NOT NULL THEN (SAFE_DIVIDE(cal_accuracy_rate, 90) * 100) ELSE NULL END AS cal_ptg,
      CASE WHEN eval_accuracy_rate IS NOT NULL THEN (SAFE_DIVIDE(eval_accuracy_rate, 90) * 100) ELSE NULL END AS eval_ptg,
      CASE WHEN dispute_rate IS NOT NULL THEN (SAFE_DIVIDE(100 - dispute_rate, 0.975)) ELSE NULL END AS dispute_ptg
    FROM stage_2
  ),

  -- STEP 7: SCORE MAPPING (Points 1-4)
  stage_4 AS (
    SELECT *,
      SAFE_DIVIDE(COALESCE(cal_ptg, 0) + COALESCE(dispute_ptg, 0) + COALESCE(eval_ptg, 0), accuracy_applicable_components_count) as accuracy_pct_to_goal,
      CASE WHEN qa_actual >= 95 THEN 4 WHEN qa_actual >= 90 THEN 3 WHEN qa_actual >= 85 THEN 2 WHEN qa_actual IS NOT NULL THEN 1 ELSE NULL END as qa_category_score,
      CASE WHEN voc_actual >= 90 THEN 4 WHEN voc_actual >= 80 THEN 3 WHEN voc_actual >= 70 THEN 2 WHEN voc_actual IS NOT NULL THEN 1 ELSE NULL END as voc_category_score,
      CASE WHEN td_actual >= 108.33 THEN 4 WHEN td_actual >= 91.67 THEN 3 WHEN td_actual >= 66.67 THEN 2 WHEN td_actual IS NOT NULL THEN 1 ELSE NULL END as td_category_score,
      CASE WHEN ppd_actual >= 108.33 THEN 4 WHEN ppd_actual >= 91.67 THEN 3 WHEN ppd_actual >= 66.67 THEN 2 WHEN ppd_actual IS NOT NULL THEN 1 ELSE NULL END as ppd_category_score
    FROM stage_3
  ),

  -- STEP 8: FINAL CATEGORY SCORING
  stage_5 AS (
    SELECT *,
      CASE WHEN productivity_pct_to_goal >= 110 THEN 4 WHEN productivity_pct_to_goal >= 100 THEN 3 WHEN productivity_pct_to_goal >= 90 THEN 2 WHEN productivity_pct_to_goal IS NOT NULL THEN 1 ELSE NULL END as productivity_category_score,
      CASE WHEN accuracy_pct_to_goal >= 98 THEN 4 WHEN accuracy_pct_to_goal >= 95 THEN 3 WHEN accuracy_pct_to_goal >= 92 THEN 2 WHEN accuracy_pct_to_goal IS NOT NULL THEN 1 ELSE NULL END as accuracy_category_score
    FROM stage_4
  ),

  -- STEP 9: GLOBAL SCORE CONSOLIDATION
  final_scores AS (
    SELECT *,
      SAFE_DIVIDE(
        COALESCE(qa_category_score, 0) + COALESCE(voc_category_score, 0) + COALESCE(td_category_score, 0) + 
        COALESCE(ppd_category_score, 0) + COALESCE(productivity_category_score, 0) + COALESCE(accuracy_category_score, 0),
        (CASE WHEN qa_category_score IS NOT NULL THEN 1 ELSE 0 END) +
        (CASE WHEN voc_category_score IS NOT NULL THEN 1 ELSE 0 END) +
        (CASE WHEN td_category_score IS NOT NULL THEN 1 ELSE 0 END) +
        (CASE WHEN ppd_category_score IS NOT NULL THEN 1 ELSE 0 END) +
        (CASE WHEN productivity_category_score IS NOT NULL THEN 1 ELSE 0 END) +
        (CASE WHEN accuracy_category_score IS NOT NULL THEN 1 ELSE 0 END)
      ) as final_score
    FROM stage_5
  )

  -- FINAL OUTPUT SELECTION & PRIMARY KEY GENERATION
  SELECT 
    LOWER(CONCAT(CAST(employee_wdid AS STRING), '_', CAST(report_year AS STRING), '_', LOWER(report_month), '_', REPLACE(LOWER(program), ' ', '_'))) AS pk_qa,
    LOWER(CONCAT(CAST(tl_wdid AS STRING), '_', CAST(report_year AS STRING), '_', LOWER(report_month), '_', REPLACE(LOWER(program), ' ', '_'))) AS pk_tl,
    LOWER(CONCAT(CAST(tl_manager_wdid AS STRING), '_', CAST(report_year AS STRING), '_', LOWER(report_month), '_', REPLACE(LOWER(program), ' ', '_'))) AS pk_manager,
    LOWER(CONCAT(CAST(tl_sr_manager_wdid AS STRING), '_', CAST(report_year AS STRING), '_', LOWER(report_month), '_', REPLACE(LOWER(program), ' ', '_'))) AS pk_sr_manager,

    employee_wdid, employee_name, report_month, report_year, program, country, region,
    tl_wdid, tl_name, tl_manager_name, tl_manager_wdid, tl_sr_manager_name, tl_sr_manager_wdid,

    qa_actual, qa_target_program, qa_weight, (SAFE_DIVIDE(qa_actual, qa_target_program) * qa_weight) as qa_weighted_score, qa_category_score, qa_notes, qa_applicable, qa_exempt,
    voc_actual, voc_target_program, voc_weight, (SAFE_DIVIDE(voc_actual, voc_target_program) * voc_weight) as voc_weighted_score, voc_category_score, voc_notes, voc_applicable, voc_exempt,
    td_actual, 100.0 as td_target_program, td_weight, (SAFE_DIVIDE(td_actual, 100.0) * td_weight) as td_weighted_score, td_category_score, td_notes, td_applicable, td_exempt,
    ppd_actual, 100.0 as ppd_target_program, ppd_weight, (SAFE_DIVIDE(ppd_actual, 100.0) * ppd_weight) as ppd_weighted_score, ppd_category_score, ppd_notes, ppd_applicable, ppd_exempt,
    
    productivity_pct_to_goal, (SAFE_DIVIDE(productivity_pct_to_goal, 100) * prod_weight) as productivity_weighted_score, productivity_category_score, prod_notes, prod_applicable, prod_exempt,
    accuracy_pct_to_goal, (SAFE_DIVIDE(accuracy_pct_to_goal, 100) * accuracy_weight) as accuracy_earned_weight, accuracy_category_score, acc_notes, acc_applicable, acc_exempt,

    final_score,
    CASE 
      WHEN final_score >= 3.50 THEN 'Excelling'
      WHEN final_score >= 3.00 THEN 'Achieving'
      WHEN final_score >= 2.00 THEN 'Developing'
      WHEN final_score IS NOT NULL THEN 'Improvement Required'
      ELSE NULL 
    END AS final_category

  FROM final_scores;

END;
