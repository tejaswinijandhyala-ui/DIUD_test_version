"""
dashboard_queries.py
====================

KEY INSIGHT — each cohort has its OWN fiscal year label:
  1% cohort: create_fy  = toYear(create_date)        + if(month>=4,1,0)
  5% cohort: create_fy  = toYear(became_5_deal_date)  + if(month>=4,1,0)
 10% cohort: create_fy  = toYear(became_10_deal_date) + if(month>=4,1,0)
 20% cohort: qual_fy    = toYear(became_20_deal_date)  + if(month>=4,1,0)

BASE_CTE WHERE (applied to all rows, inherited by every query):
  - pipeline = 'default'
  - deal_type NOT IN ('Partner-Led SMB')
  - deal_id IN (gs_deal_ids_hs)
  - is_this_a_deal_with_inception: NOT filtered (commented out in all 3 dashboard sqls)
"""


# ── 1% cohort ─────────────────────────────────────────────────────────────────
# Dashboard 1% SQL:
#   WHERE create_date >= '2025-04-01' AND deal_stage IN (all stages)
#   final WHERE: create_fy >= 2026  (create_fy = toYear(create_date)+if(month>=4,1,0))
# Extra condition from 1_creation CTE:
#   (create_date <> became_5_deal_date OR deal_stage = '1% - IQM Scheduled')
#   (create_date <> became_10_deal_date OR deal_stage = '1% - IQM Scheduled')

_DEFAULT_FY     = 2027          # shown when no FY filter is selected
_DEFAULT_FY_START = '2026-04-01'  # FY27 starts April 1 2026
_FY26_START     = '2025-04-01'
_FY27_START     = '2026-04-01'

def _make_pct_filters(fy_list: list):
    """Return PCT filter strings for a given list of fiscal years."""
    if len(fy_list) == 1:
        fy       = fy_list[0]
        fy_start = _FY26_START if fy == 2026 else _FY27_START
        fy_clause_1  = f"create_fy >= {fy}"
        fy_clause_5  = f"fy_5 = {fy}"
        fy_clause_10 = f"fy_10 = {fy}"
        fy_clause_20 = f"qual_fy = {fy}"
    else:
        # [2026, 2027] — All FY or both checked
        in_clause    = ", ".join(str(f) for f in sorted(fy_list))
        fy_start     = _FY26_START
        fy_clause_1  = f"create_fy IN ({in_clause})"
        fy_clause_5  = f"fy_5 IN ({in_clause})"
        fy_clause_10 = f"fy_10 IN ({in_clause})"
        fy_clause_20 = f"qual_fy IN ({in_clause})"

    _1PCT = f"""
    create_date >= '{fy_start}'
    AND create_date <> '1900-01-01'
    AND {fy_clause_1}
    AND deal_stage IN (
        '1% - IQM Scheduled',
        '5% - IQM Held',
        '10% - Discovery',
        '20% - Solution',
        '30% - Proof',
        '40% - Proposal',
        '60% - Price Negotiation',
        '75% - Contract Review',
        '90% - Deal Desk Review',
        'Closed Won',
        'Closed Lost',
        'Didn''t Qualify',
        'Prospect Disengaged',
        'Deal on Hold'
    )
"""
    _5PCT = f"""
    became_5_deal_date >= '{fy_start}'
    AND became_5_deal_date <> '1900-01-01'
    AND {fy_clause_5}
    AND deal_stage IN (
        '5% - IQM Held',
        '10% - Discovery',
        '20% - Solution',
        '30% - Proof',
        '40% - Proposal',
        '60% - Price Negotiation',
        '75% - Contract Review',
        '90% - Deal Desk Review',
        'Closed Won',
        'Closed Lost',
        'Didn''t Qualify',
        'Prospect Disengaged',
        'Deal on Hold'
    )
"""
    _10PCT = f"""
    became_10_deal_date >= '{fy_start}'
    AND became_10_deal_date <> '1900-01-01'
    AND {fy_clause_10}
    AND deal_stage IN (
        '10% - Discovery',
        '20% - Solution',
        '30% - Proof',
        '40% - Proposal',
        '60% - Price Negotiation',
        '75% - Contract Review',
        '90% - Deal Desk Review',
        'Closed Won',
        'Closed Lost',
        'Didn''t Qualify',
        'Prospect Disengaged',
        'Deal on Hold'
    )
"""
    _20PCT = f"""
    became_20_deal_date >= '{fy_start}'
    AND became_20_deal_date <> '1900-01-01'
    AND {fy_clause_20}
    AND deal_stage IN (
        '20% - Solution',
        '30% - Proof',
        '40% - Proposal',
        '60% - Price Negotiation',
        '75% - Contract Review',
        '90% - Deal Desk Review',
        'Closed Won',
        'Closed Lost',
        'Didn''t Qualify',
        'Prospect Disengaged',
        'Deal on Hold'
    )
"""
    _20PCT_ACTIVE = f"""
    became_20_deal_date >= '{fy_start}'
    AND became_20_deal_date <> '1900-01-01'
    AND {fy_clause_20}
    AND deal_stage IN (
        '20% - Solution',
        '30% - Proof',
        '40% - Proposal',
        '60% - Price Negotiation',
        '75% - Contract Review'
    )
"""
    return _1PCT, _5PCT, _10PCT, _20PCT, _20PCT_ACTIVE


# was: _make_pct_filters(_DEFAULT_FY)
_1PCT_FILTER, _5PCT_FILTER, _10PCT_FILTER, _20PCT_FILTER, _20PCT_ACTIVE_FILTER = _make_pct_filters([_DEFAULT_FY])

# ── Current quarter/month helpers ─────────────────────────────────────────────
_CURRENT_QUARTER = """CASE
    WHEN toMonth(CURRENT_DATE()) IN (4,5,6)    THEN 'Q1'
    WHEN toMonth(CURRENT_DATE()) IN (7,8,9)    THEN 'Q2'
    WHEN toMonth(CURRENT_DATE()) IN (10,11,12) THEN 'Q3'
    WHEN toMonth(CURRENT_DATE()) IN (1,2,3)    THEN 'Q4'
END"""

_CURRENT_MONTH = "LEFT(formatDateTime(CURRENT_DATE(), '%M'), 3)"


# =============================================================================
# SINGLE BASE CTE
# Adds fy_5 and fy_10 columns — fiscal year computed from became_5/10_deal_date,
# matching exactly how your dashboard 5% and 10% SQL compute create_fy.
# =============================================================================

BASE_CTE = """
WITH pipe_gen AS (
    SELECT
        toInt64(d.deal_id)                                                       AS deal_id,
        d.deal_name,
        concat(o.firstName, ' ', o.lastName)                                     AS deal_owner_name,
        COALESCE(t.name, 'N/A')                                                  AS team,
        d.deal_url,
        d.country,

        CAST(LEFT(coalesce(d.create_date,         '1900-01-01'), 10) AS DATE)    AS create_date,
        CAST(LEFT(coalesce(d.close_date,          '1900-01-01'), 10) AS DATE)    AS close_date,
        CAST(LEFT(coalesce(d.became_5_deal_date,  '1900-01-01'), 10) AS DATE)    AS became_5_deal_date,
        CAST(LEFT(coalesce(d.became_10_deal_date, '1900-01-01'), 10) AS DATE)    AS became_10_deal_date,
        CAST(LEFT(coalesce(d.became_20_deal_date, '1900-01-01'), 10) AS DATE)    AS became_20_deal_date,
        CAST(LEFT(coalesce(d.became_30_deal_date, '1900-01-01'), 10) AS DATE)    AS became_30_deal_date,
        CAST(LEFT(coalesce(d.became_40_deal_date, '1900-01-01'), 10) AS DATE)    AS became_40_deal_date,
        CAST(LEFT(coalesce(d.became_60_deal_date, '1900-01-01'), 10) AS DATE)    AS became_60_deal_date,
        CAST(LEFT(coalesce(d.became_75_deal_date, '1900-01-01'), 10) AS DATE)    AS became_75_deal_date,
        CAST(LEFT(coalesce(d.last_contacted,      '1900-01-01'), 10) AS DATE)    AS last_contacted,

        d.deal_stage,
        d.pipeline,
        CASE WHEN d.deal_type IS NULL THEN 'Not Assigned' ELSE d.deal_type END   AS deal_type,

        d.amount,
        d.forecast_amount,
        d.forecast_probability,
        d.management_forecast,

        CASE
            WHEN d.region = 'india___sea' THEN 'ISEA'
            WHEN d.region = 'Africa'      THEN 'Middle East'
            WHEN d.region = 'japac'       THEN 'JAPAC'
            ELSE d.region
        END                                                                       AS region,

        CASE
            WHEN d.deal_source_rollup IN ('Executive Outreach', 'Investor')      THEN 'Executive Outreach'
            WHEN d.deal_source_rollup IN ('Marketing', 'Customer Success',
                                          'AE Outbound', 'Inception',
                                          'Hyperscaler')                         THEN d.deal_source_rollup
            WHEN d.deal_source_rollup IN ('BDR Outbound')                        THEN 'BDR'
            WHEN d.deal_source_rollup IN ('Partner')                             THEN 'Partner - Non Hyperscaler'
            ELSE 'Other'
        END                                                                       AS deal_source_rollup,

        CASE WHEN d.ai_for_x IS NULL THEN 'N/A' ELSE d.ai_for_x END             AS ai_for_x,

        CASE
            WHEN d.kore_primary_industry IN ('Financial Services',
                                             'Banking', 'Insurance')             THEN 'Financial Services'
            WHEN d.kore_primary_industry IN ('Manufacturing Discreet',
                                             'Manufacturing Process', 'CPG')     THEN 'Manufacturing'
            WHEN d.kore_primary_industry IN ('Hi-Tech',
                                             'Telecom / Media / Entertainment')  THEN 'TMT'
            WHEN d.kore_primary_industry IS NULL
              OR d.kore_primary_industry IN ('Business Services', 'Government',
                                             'Energy & Utilities', 'Education',
                                             'Restaurants', 'null', 'Energy')    THEN 'Other'
            ELSE d.kore_primary_industry
        END                                                                       AS kore_primary_industry,

        CASE
            WHEN d.is_there_a_confirmation_of_budget = 'Yes'
             AND d.who_is_the_decision_maker IS NOT NULL
             AND d.use_case IS NOT NULL
             AND d.what_is_the_estimated_timeline IS NOT NULL THEN 'Yes'
            ELSE 'No'
        END                                                                       AS BANT,


        CASE WHEN d.is_this_a_deal_with_inception = 'Yes'
             THEN 'Yes' ELSE 'No' END                                            AS is_inception_deal,

        CASE
            WHEN d.deal_stage = '1% - IQM Scheduled'      THEN 7
            WHEN d.deal_stage = '5% - IQM Held'           THEN 21
            WHEN d.deal_stage = '10% - Discovery'         THEN 28
            WHEN d.deal_stage = '20% - Solution'          THEN 41
            WHEN d.deal_stage = '30% - Proof'             THEN 15
            WHEN d.deal_stage = '40% - Proposal'          THEN 29
            WHEN d.deal_stage = '60% - Price Negotiation' THEN 27
            WHEN d.deal_stage = '75% - Contract Review'   THEN 34
            ELSE NULL
        END                                                                       AS avg_days_benchmark,

        CASE
            WHEN d.deal_stage IN ('Prospect Disengaged', 'Closed Lost',
                                  'Didn''t Qualify', '90% - Deal Desk Review',
                                  'Closed Won')                                  THEN NULL
            WHEN d.deal_stage = '1% - IQM Scheduled'
                AND CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '5% - IQM Held'
                AND CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '10% - Discovery'
                AND CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '20% - Solution'
                AND CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '30% - Proof'
                AND CAST(LEFT(coalesce(d.became_30_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_30_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '40% - Proposal'
                AND CAST(LEFT(coalesce(d.became_40_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_40_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '60% - Price Negotiation'
                AND CAST(LEFT(coalesce(d.became_60_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_60_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            WHEN d.deal_stage = '75% - Contract Review'
                AND CAST(LEFT(coalesce(d.became_75_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_75_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE())
            ELSE NULL
        END                                                                       AS days_in_current_stage,

        CASE
            WHEN d.deal_stage = '1% - IQM Scheduled'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) <  7*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) <  7*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '5% - IQM Held'
                AND CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 21*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 21*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '10% - Discovery'
                AND CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 28*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 28*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '20% - Solution'
                AND CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 41*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 41*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '30% - Proof'
                AND CAST(LEFT(coalesce(d.became_30_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_30_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 15*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_30_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 15*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '40% - Proposal'
                AND CAST(LEFT(coalesce(d.became_40_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_40_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 29*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_40_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 29*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '60% - Price Negotiation'
                AND CAST(LEFT(coalesce(d.became_60_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_60_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 27*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_60_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 27*2   THEN 'Yellow'
                    ELSE 'Red' END
            WHEN d.deal_stage = '75% - Contract Review'
                AND CAST(LEFT(coalesce(d.became_75_deal_date,'1900-01-01'),10) AS DATE) <> '1900-01-01'
                THEN CASE
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_75_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 34*1.5 THEN 'Green'
                    WHEN DATE_DIFF('Day', CAST(LEFT(coalesce(d.became_75_deal_date,'1900-01-01'),10) AS DATE), CURRENT_DATE()) < 34*2   THEN 'Yellow'
                    ELSE 'Red' END
            ELSE NULL
        END                                                                       AS deal_health,

        toYear(CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE))
            + if(toMonth(CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)
                                                                                  AS create_fy,
        CASE
            WHEN toMonth(CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE)) IN (4,5,6)    THEN 'Q1'
            WHEN toMonth(CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE)) IN (7,8,9)    THEN 'Q2'
            WHEN toMonth(CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE)) IN (10,11,12) THEN 'Q3'
            WHEN toMonth(CAST(LEFT(coalesce(d.create_date,'1900-01-01'),10) AS DATE)) IN (1,2,3)    THEN 'Q4'
        END                                                                       AS create_quarter,

        -- Matches your 5_creation CTE: create_fy = toYear(became_5_deal_date)+if(month>=4,1,0)
        toYear(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE))
            + if(toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)
                                                                                  AS fy_5,
        CASE
            WHEN toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) IN (4,5,6)    THEN 'Q1'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) IN (7,8,9)    THEN 'Q2'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) IN (10,11,12) THEN 'Q3'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) IN (1,2,3)    THEN 'Q4'
        END                                                                       AS quarter_5,

        -- ── FY from became_10_deal_date (used by 10% cohort) ─────────────────
        -- Matches your 10_creation CTE: create_fy = toYear(became_10_deal_date)+if(month>=4,1,0)
        toYear(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE))
            + if(toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)
                                                                                  AS fy_10,
        CASE
            WHEN toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) IN (4,5,6)    THEN 'Q1'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) IN (7,8,9)    THEN 'Q2'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) IN (10,11,12) THEN 'Q3'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) IN (1,2,3)    THEN 'Q4'
        END                                                                       AS quarter_10,

        -- ── FY from became_20_deal_date (used by 20% cohort) ─────────────────
        toYear(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE))
            + if(toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)
                                                                                  AS qual_fy,
        CASE
            WHEN toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) IN (4,5,6)    THEN 'Q1'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) IN (7,8,9)    THEN 'Q2'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) IN (10,11,12) THEN 'Q3'
            WHEN toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) IN (1,2,3)    THEN 'Q4'
        END                                                                       AS qual_quarter,

        -- ── FY from close_date ────────────────────────────────────────────────
        toYear(CAST(LEFT(coalesce(d.close_date,'1900-01-01'),10) AS DATE))
            + if(toMonth(CAST(LEFT(coalesce(d.close_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)
                                                                                  AS close_fy,
        CASE
            WHEN toMonth(CAST(LEFT(coalesce(d.close_date,'1900-01-01'),10) AS DATE)) IN (4,5,6)    THEN 'Q1'
            WHEN toMonth(CAST(LEFT(coalesce(d.close_date,'1900-01-01'),10) AS DATE)) IN (7,8,9)    THEN 'Q2'
            WHEN toMonth(CAST(LEFT(coalesce(d.close_date,'1900-01-01'),10) AS DATE)) IN (10,11,12) THEN 'Q3'
            WHEN toMonth(CAST(LEFT(coalesce(d.close_date,'1900-01-01'),10) AS DATE)) IN (1,2,3)    THEN 'Q4'
        END                                                                       AS close_quarter

    FROM hs_analytics.deals d FINAL
    LEFT JOIN hs_analytics.owners o FINAL
           ON d.deal_owner = CAST(o.id AS VARCHAR)
    LEFT JOIN kore_ai_hubspot.gs_Teams t
           ON d.hubspot_team = t.team_id
    WHERE d.pipeline = 'default'
      AND CASE WHEN d.deal_type IS NULL 
      THEN 'Not Assigned' 
      ELSE d.deal_type 
      END NOT IN ('Partner-Led SMB')
      AND d.deal_stage IN (
            '1% - IQM Scheduled',
            '5% - IQM Held',
            '10% - Discovery',
            '20% - Solution',
            '30% - Proof',
            '40% - Proposal',
            '60% - Price Negotiation',
            '75% - Contract Review',
            '90% - Deal Desk Review',
            'Closed Won',
            'Closed Lost',
            'Didn''t Qualify',
            'Prospect Disengaged',
            'Deal on Hold'
      )
      AND toInt64(d.deal_id) IN (
            SELECT DISTINCT toInt64(deal_id_hs)
            FROM kore_ai_hubspot.gs_deal_ids_hs
          )
)
"""

# =============================================================================
# FILTERED BASE CTE BUILDER
# =============================================================================

def _to_sql_in_list_str(raw: str) -> list:
    """
    Splits a comma-separated filter string into a list of non-empty, trimmed values,
    stripping the sentinel value 'ALL'.  Returns [] if nothing remains.
    """
    return [v.strip() for v in raw.split(",") if v.strip() and v.strip().upper() != "ALL"]


def _build_in_clause_str(values: list) -> str:
    """Build a SQL IN (...) clause for string columns (lower-cased comparison)."""
    quoted = ", ".join(f"lower('{v.replace(chr(39), chr(39)*2)}')" for v in values)
    return f"({quoted})"


def _build_in_clause_int(values: list) -> str:
    """Build a SQL IN (...) clause for integer columns (FY)."""
    ints = ", ".join(v for v in values if v.isdigit())
    return f"({ints})"


def build_filtered_base_cte(filters: dict) -> str:
    """
    Takes a filters dict with keys:
      region (str), deal_source (str), fy (str/int),
      ai_for_x (str), industry (str), stage (str: '5','10','20', or comma-separated)
    Supports multiple comma-separated values per filter (sent by multi-select UI).
    Returns a modified BASE_CTE string with extra WHERE conditions injected.

    IMPORTANT: All extra_conditions must be plain predicates (no leading AND).
    They are joined with AND and appended to the existing WHERE clause.

    NOTE on stage filter: The _xPCT_FILTER snippets reference computed aliases
    (qual_fy, fy_5, fy_10) that are not available inside the pipe_gen CTE WHERE.
    We therefore rewrite them using the underlying base expressions instead.
    """
    extra_conditions = []

    # ── Region ────────────────────────────────────────────────────────────────
    region_vals = _to_sql_in_list_str((filters.get("region") or "").strip())
    if region_vals:
        if len(region_vals) == 1:
            safe = region_vals[0].replace("'", "''")
            extra_conditions.append(f"lower(d.region) = lower('{safe}')")
        else:
            in_clause = _build_in_clause_str(region_vals)
            extra_conditions.append(f"lower(d.region) IN {in_clause}")

    # ── Deal Source ───────────────────────────────────────────────────────────
    # Filter on RAW d.deal_source_rollup values (before pipe_gen CASE rolls them up)
    _ACTUALS_SOURCE_REVERSE_MAP = {
        "BDR":                       ["BDR Outbound", "BDR"],
        "Marketing":                 ["Marketing"],
        "Partner - Non Hyperscaler": ["Partner"],
        "Hyperscaler":               ["Hyperscaler"],
        "Customer Success":          ["Customer Success"],
        "Executive Outreach":        ["Executive Outreach", "Investor"],
        "Inception":                 ["Inception"],
    }
    source_vals = _to_sql_in_list_str((filters.get("deal_source") or "").strip())
    if source_vals:
        expanded = []
        for v in source_vals:
            expanded.extend(_ACTUALS_SOURCE_REVERSE_MAP.get(v, [v]))
        quoted = ", ".join(f"'{r.replace(chr(39), chr(39)*2)}'" for r in expanded)
        extra_conditions.append(f"d.deal_source_rollup IN ({quoted})")

    # ── Fiscal Year ───────────────────────────────────────────────────────────
    # NOTE: fy_5, fy_10, qual_fy, create_fy are aliases computed in the SELECT,
    # so we cannot reference them in the WHERE. Rewrite inline using base columns.
    fy_vals = _to_sql_in_list_str((filters.get("fy") or "").strip())
    fy_vals = [v for v in fy_vals if v.isdigit()]
    if fy_vals:
        def _fy_expr(col: str) -> str:
            """Inline fiscal-year expression for a raw date column."""
            return (
                f"(toYear(CAST(LEFT(coalesce(d.{col},'1900-01-01'),10) AS DATE))"
                f" + if(toMonth(CAST(LEFT(coalesce(d.{col},'1900-01-01'),10) AS DATE)) >= 4, 1, 0))"
            )
        fy_create   = _fy_expr("create_date")
        fy_5_expr   = _fy_expr("became_5_deal_date")
        fy_10_expr  = _fy_expr("became_10_deal_date")
        fy_20_expr  = _fy_expr("became_20_deal_date")
        if len(fy_vals) == 1:
            fy_int = fy_vals[0]
            extra_conditions.append(
                f"({fy_create} = {fy_int}"
                f" OR {fy_5_expr} = {fy_int}"
                f" OR {fy_10_expr} = {fy_int}"
                f" OR {fy_20_expr} = {fy_int})"
            )
        else:
            in_list = ", ".join(fy_vals)
            extra_conditions.append(
                f"({fy_create} IN ({in_list})"
                f" OR {fy_5_expr} IN ({in_list})"
                f" OR {fy_10_expr} IN ({in_list})"
                f" OR {fy_20_expr} IN ({in_list}))"
            )

    # ── AI for X ──────────────────────────────────────────────────────────────
    ai_vals = _to_sql_in_list_str((filters.get("ai_for_x") or "").strip())
    if ai_vals:
        if len(ai_vals) == 1:
            safe = ai_vals[0].replace("'", "''")
            extra_conditions.append(f"lower(d.ai_for_x) = lower('{safe}')")
        else:
            in_clause = _build_in_clause_str(ai_vals)
            extra_conditions.append(f"lower(d.ai_for_x) IN {in_clause}")

    # ── Industry ──────────────────────────────────────────────────────────────
    # Must filter on RAW d.kore_primary_industry values, not the rolled-up label.
    # The CASE in pipe_gen runs AFTER the WHERE, so filtering on rollup label = 0 rows.
    _INDUSTRY_REVERSE_MAP = {
        "Financial Services":        ["Financial Services", "Banking", "Insurance"],
        "Manufacturing":             ["Manufacturing Discreet", "Manufacturing Process", "CPG"],
        "TMT":                       ["Hi-Tech", "Telecom / Media / Entertainment"],
        "Other":                     ["Business Services", "Government", "Energy & Utilities",
                                      "Education", "Restaurants", "null", "Energy"],
        "Retail":                    ["Retail"],
        "Healthcare":                ["Healthcare"],
        "Travel & Transportation":   ["Travel & Transportation"],
        "Healthcare Payer":          ["Healthcare Payer"],
        "Healthcare Life Sciences":  ["Healthcare Life Sciences"],
    }
    industry_vals = _to_sql_in_list_str((filters.get("industry") or "").strip())
    if industry_vals:
        expanded = []
        for v in industry_vals:
            expanded.extend(_INDUSTRY_REVERSE_MAP.get(v, [v]))
        quoted = ", ".join(f"'{r.replace(chr(39), chr(39)*2)}'" for r in expanded)
        extra_conditions.append(f"d.kore_primary_industry IN ({quoted})")

    # ── Stage ─────────────────────────────────────────────────────────────────
    # Supports single ("5") or multi ("5,10") selection.
    # When multiple stages selected, results are the UNION (OR) of each cohort filter.
    # NOTE: _xPCT_FILTER snippets reference computed aliases (qual_fy, fy_5, fy_10)
    # which are not available in the pipe_gen CTE WHERE clause.
    # We rewrite them inline using the underlying raw date columns instead.
    # Determine which FY to use for stage-filter inline expressions
    _selected_fy_vals = _to_sql_in_list_str((filters.get("fy") or "").strip())
    _selected_fy_vals = [v for v in _selected_fy_vals if v.isdigit()]

    if len(_selected_fy_vals) == 1:
        _stage_fy       = int(_selected_fy_vals[0])
        _stage_fy_start = _FY26_START if _stage_fy == 2026 else _DEFAULT_FY_START
        _fy5_clause     = f"(toYear(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) + if(toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)) = {_stage_fy}"
        _fy10_clause    = f"(toYear(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) + if(toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)) = {_stage_fy}"
        _fy20_clause    = f"(toYear(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) + if(toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)) = {_stage_fy}"
    else:
        # All FY or both selected — include both years
        _stage_fy_start = _FY26_START
        _fy5_clause     = f"(toYear(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) + if(toMonth(CAST(LEFT(coalesce(d.became_5_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)) IN (2026, 2027)"
        _fy10_clause    = f"(toYear(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) + if(toMonth(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)) IN (2026, 2027)"
        _fy20_clause    = f"(toYear(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) + if(toMonth(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE)) >= 4, 1, 0)) IN (2026, 2027)"

    _FY_INLINE = {
        "5":  (
            f"d.became_5_deal_date >= '{_stage_fy_start}'"
            " AND d.became_5_deal_date <> '1900-01-01'"
            f" AND {_fy5_clause}"
            " AND d.deal_stage IN ("
            "  '5% - IQM Held','10% - Discovery','20% - Solution','30% - Proof',"
            "  '40% - Proposal','60% - Price Negotiation','75% - Contract Review',"
            "  '90% - Deal Desk Review','Closed Won','Closed Lost',"
            "  'Didn''t Qualify','Prospect Disengaged','Deal on Hold')"
        ),
        "10": (
            f"d.became_10_deal_date >= '{_stage_fy_start}'"
            " AND d.became_10_deal_date <> '1900-01-01'"
            " AND (toYear(CAST(LEFT(coalesce(d.became_10_deal_date,'1900-01-01'),10) AS DATE))"
            f" AND {_fy10_clause}"
            " AND d.deal_stage IN ("
            "  '10% - Discovery','20% - Solution','30% - Proof','40% - Proposal',"
            "  '60% - Price Negotiation','75% - Contract Review',"
            "  '90% - Deal Desk Review','Closed Won','Closed Lost',"
            "  'Didn''t Qualify','Prospect Disengaged','Deal on Hold')"
        ),
        "20": (
            f"d.became_20_deal_date >= '{_stage_fy_start}'"
            " AND d.became_20_deal_date <> '1900-01-01'"
            " AND (toYear(CAST(LEFT(coalesce(d.became_20_deal_date,'1900-01-01'),10) AS DATE))"
            f" AND {_fy20_clause}"
            " AND d.deal_stage IN ("
            "  '20% - Solution','30% - Proof','40% - Proposal',"
            "  '60% - Price Negotiation','75% - Contract Review',"
            "  '90% - Deal Desk Review','Closed Won','Closed Lost',"
            "  'Didn''t Qualify','Prospect Disengaged','Deal on Hold')"
        ),
    }
    stage_vals = _to_sql_in_list_str((filters.get("stage") or "").strip())
    stage_clauses = []
    for s in stage_vals:
        if s in _FY_INLINE:
            stage_clauses.append(f"({_FY_INLINE[s]})")
    if stage_clauses:
        combined = "\n          OR ".join(stage_clauses)
        extra_conditions.append(f"(\n          {combined}\n        )")

    if not extra_conditions:
        return BASE_CTE  # no filters — return as-is

    # Inject before the closing `)` of the CTE.
    # All items in extra_conditions are plain predicates — join them with AND.
    extra_sql = "\n      AND " + "\n      AND ".join(extra_conditions)
    return BASE_CTE.replace(
        "          )\n)",
        "          )" + extra_sql + "\n)"
    )



def build_target_filters(filters: dict) -> str:
    conditions = []

    # ── Fiscal Year ───────────────────────────────────────────────────────────
    fy_vals = _to_sql_in_list_str((filters.get("fy") or "").strip())
    fy_vals = [v for v in fy_vals if v.isdigit()]
    if fy_vals:
        if len(fy_vals) == 1:
            conditions.append(f"CAST(fy AS INT) = {fy_vals[0]}")
        else:
            in_clause = _build_in_clause_int(fy_vals)
            conditions.append(f"CAST(fy AS INT) IN {in_clause}")
    else:
        # "All FY" selected — include both FY26 and FY27 targets
        conditions.append(f"CAST(fy AS INT) IN (2026, 2027)")

    # ── Region ────────────────────────────────────────────────────────────────
    region_vals = _to_sql_in_list_str((filters.get("region") or "").strip())
    if region_vals:
        if len(region_vals) == 1:
            safe = region_vals[0].replace("'", "''")
            conditions.append(f"lower(region) = lower('{safe}')")
        else:
            in_clause = _build_in_clause_str(region_vals)
            conditions.append(f"lower(region) IN {in_clause}")

    # ── Deal Source ───────────────────────────────────────────────────────────
    # Reverse-map: pipe_gen rollup label → raw values in gs_pipeline_quotas_v1.source
    _SOURCE_REVERSE_MAP = {
        "BDR":                       ["BDR Outbound", "BDR"],
        "Marketing":                 ["Marketing"],
        "Partner - Non Hyperscaler": ["Partner", "Partner - Excluding Hyperscalers", "Partner - Non Hyperscaler"],
        "Hyperscaler":               ["Hyperscaler", "Hyperscalers"],
        "Customer Success":          ["Customer Success"],
        "Executive Outreach":        ["Executive Outreach", "Investor"],
        "Inception":                 ["Inception"],
    }
    source_vals = _to_sql_in_list_str((filters.get("deal_source") or "").strip())
    if source_vals:
        expanded = []
        for v in source_vals:
            raw_list = _SOURCE_REVERSE_MAP.get(v, [v])
            expanded.extend(raw_list)
        quoted = ", ".join(f"'{r.replace(chr(39), chr(39)*2)}'" for r in expanded)
        conditions.append(f"source IN ({quoted})")

    # OPTIONAL (only if columns exist in targets table)
    month = (filters.get("month") or "").strip()
    if month:
        conditions.append(f"month = '{month}'")

    return " AND ".join(conditions) if conditions else "1=1"



def get_filtered_pipeline_metrics(client, filters: dict):
    filtered_cte = build_filtered_base_cte(filters)
    target_filters = build_target_filters(filters)

    # Determine which FY the user selected — default to _DEFAULT_FY
    # Determine which FYs the user selected
    fy_vals = _to_sql_in_list_str((filters.get("fy") or "").strip())
    fy_list = [int(v) for v in fy_vals if v.isdigit()]
    # fy_list = []          → "All FY" selected (ALL stripped) → treat as both
    # fy_list = [2027]      → FY27 only
    # fy_list = [2026]      → FY26 only
    # fy_list = [2026,2027] → both checked
    if not fy_list:
        fy_list = [2026, 2027]

    f1, f5, f10, f20, f20a = _make_pct_filters(fy_list)

    def rebuild(original_query):
        q = filtered_cte + original_query[len(BASE_CTE):]
        # Swap in the FY-correct PCT filters
        q = q.replace(_1PCT_FILTER,          f1)
        q = q.replace(_5PCT_FILTER,          f5)
        q = q.replace(_10PCT_FILTER,         f10)
        q = q.replace(_20PCT_FILTER,         f20)
        q = q.replace(_20PCT_ACTIVE_FILTER,  f20a)
        return q.replace("{TARGET_FILTERS}", target_filters)

    def run(sql, label):
        try:
            r = client.query(sql)
            rows = [dict(zip(r.column_names, row)) for row in r.result_rows]
            print(f"✅ [filtered] {label}: {len(rows)} row(s)")
            return rows
        except Exception as e:
            print(f"❌ [filtered] {label} failed: {e}")
            return []

    return {
        "funnel":           (run(rebuild(QUERY_FUNNEL_OVERVIEW),           "Funnel") or [{}])[0],
        "quarterly_trend":   run(rebuild(QUERY_QUARTERLY_TREND),           "Quarterly Trend"),
        "region_source":     run(rebuild(QUERY_REGION_SOURCE_PERFORMANCE), "Region/Source"),
        "stage_velocity":    run(rebuild(QUERY_STAGE_VELOCITY),            "Stage Velocity"),
        "industry_product":  run(rebuild(QUERY_INDUSTRY_PRODUCT),          "Industry/Product"),
        "period_attainment":(run(rebuild(QUERY_PERIOD_ATTAINMENT),         "MTD/QTD/YTD") or [{}])[0],
        "deals_to_watch":    run(rebuild(QUERY_DEALS_TO_WATCH),            "Deals to Watch"),
        "funnel_conversions": (run(rebuild(_make_funnel_conversions(fy_list)), "Funnel Conversions") or [{}])[0],
        "won_lost_deals": run(QUERY_WON_LOST_DEALS, "Won/Lost Deals")
    }
    

# =============================================================================
# QUERY 1 — SLIDE 2: Funnel counts + pipeline values
# =============================================================================
QUERY_FUNNEL_OVERVIEW = BASE_CTE + f"""
SELECT
    countDistinctIf(deal_id, {_1PCT_FILTER})                                      AS cnt_1pct,

    countDistinctIf(deal_id, {_5PCT_FILTER})                                      AS cnt_5pct,
    sumIf(amount,            {_5PCT_FILTER})                                      AS amt_5pct,

    countDistinctIf(deal_id, {_10PCT_FILTER})                                     AS cnt_10pct,
    sumIf(amount,            {_10PCT_FILTER})                                     AS amt_10pct,

    countDistinctIf(deal_id, {_20PCT_FILTER})                                     AS cnt_20pct,
    sumIf(amount,            {_20PCT_FILTER})                                     AS amt_20pct,

    countDistinctIf(deal_id, {_20PCT_FILTER}
        AND deal_stage IN ('Closed Won', '90% - Deal Desk Review'))               AS cnt_closed_won,
    sumIf(amount,            {_20PCT_FILTER}
        AND deal_stage IN ('Closed Won', '90% - Deal Desk Review'))               AS amt_closed_won,

    countDistinctIf(deal_id, {_1PCT_FILTER}
        AND deal_stage IN (
            'Closed Lost', 'Prospect Disengaged', 'Didn''t Qualify'
        ))                                                                         AS cnt_fallen_out
FROM pipe_gen
"""


# New query — add after QUERY_FUNNEL_OVERVIEW

def _make_funnel_conversions(fy_list: list) -> str:
    if len(fy_list) == 1:
        fy_5_clause  = f"fy_5 = {fy_list[0]}"
        fy_10_clause = f"fy_10 = {fy_list[0]}"
        fy_20_clause = f"qual_fy = {fy_list[0]}"
    else:
        in_clause    = ", ".join(str(f) for f in sorted(fy_list))
        fy_5_clause  = f"fy_5 IN ({in_clause})"
        fy_10_clause = f"fy_10 IN ({in_clause})"
        fy_20_clause = f"qual_fy IN ({in_clause})"

    return BASE_CTE + f"""
, conv_5 AS (
    SELECT
        countDistinctIf(deal_id,
            {fy_5_clause}
            AND became_5_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
        ) AS base_5,

        countDistinctIf(deal_id,
            {fy_5_clause}
            AND became_5_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
            AND (
                became_10_deal_date <> '1900-01-01'
                OR became_20_deal_date <> '1900-01-01'
                OR became_30_deal_date <> '1900-01-01'
                OR became_40_deal_date <> '1900-01-01'
                OR became_60_deal_date <> '1900-01-01'
                OR became_75_deal_date <> '1900-01-01'
                OR deal_stage IN (
                    '10% - Discovery','20% - Solution','30% - Proof',
                    '40% - Proposal','60% - Price Negotiation',
                    '75% - Contract Review','90% - Deal Desk Review','Closed Won'
                )
            )
        ) AS conv_5_to_10,

        countDistinctIf(deal_id,
            {fy_5_clause}
            AND became_5_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
            AND (
                became_20_deal_date <> '1900-01-01'
                OR became_30_deal_date <> '1900-01-01'
                OR became_40_deal_date <> '1900-01-01'
                OR became_60_deal_date <> '1900-01-01'
                OR became_75_deal_date <> '1900-01-01'
                OR deal_stage IN (
                    '20% - Solution','30% - Proof','40% - Proposal',
                    '60% - Price Negotiation','75% - Contract Review',
                    '90% - Deal Desk Review','Closed Won'
                )
            )
        ) AS conv_5_to_20,

        countDistinctIf(deal_id,
            {fy_5_clause}
            AND became_5_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
            AND deal_stage IN ('90% - Deal Desk Review','Closed Won')
        ) AS conv_5_to_won
    FROM pipe_gen
),

conv_10 AS (
    SELECT
        countDistinctIf(deal_id,
            {fy_10_clause}
            AND became_10_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
        ) AS base_10,

        countDistinctIf(deal_id,
            {fy_10_clause}
            AND became_10_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
            AND (
                became_20_deal_date <> '1900-01-01'
                OR became_30_deal_date <> '1900-01-01'
                OR became_40_deal_date <> '1900-01-01'
                OR became_60_deal_date <> '1900-01-01'
                OR became_75_deal_date <> '1900-01-01'
                OR deal_stage IN (
                    '20% - Solution','30% - Proof','40% - Proposal',
                    '60% - Price Negotiation','75% - Contract Review',
                    '90% - Deal Desk Review','Closed Won'
                )
            )
        ) AS conv_10_to_20,

        countDistinctIf(deal_id,
            {fy_10_clause}
            AND became_10_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
            AND deal_stage IN ('90% - Deal Desk Review','Closed Won')
        ) AS conv_10_to_won
    FROM pipe_gen
),

conv_20 AS (
    SELECT
        countDistinctIf(deal_id,
            {fy_20_clause}
            AND became_20_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
        ) AS base_20,

        countDistinctIf(deal_id,
            {fy_20_clause}
            AND became_20_deal_date <> '1900-01-01'
            AND deal_type NOT IN ('Partner-Led SMB')
            AND deal_stage IN ('90% - Deal Desk Review','Closed Won')
        ) AS conv_20_to_won
    FROM pipe_gen
)

SELECT
    conv_5.base_5, conv_5.conv_5_to_10, conv_5.conv_5_to_20, conv_5.conv_5_to_won,
    conv_10.base_10, conv_10.conv_10_to_20, conv_10.conv_10_to_won,
    conv_20.base_20, conv_20.conv_20_to_won
FROM conv_5 CROSS JOIN conv_10 CROSS JOIN conv_20
"""

# Module-level default (FY27) for get_all_pipeline_metrics
# was: _make_funnel_conversions(_DEFAULT_FY)
QUERY_FUNNEL_CONVERSIONS = _make_funnel_conversions([_DEFAULT_FY])

# =============================================================================
# QUERY 2 — SLIDE 3: Quarterly actuals vs target — all 3 stages
# Note: each cohort groups by its own quarter column
# =============================================================================
QUERY_QUARTERLY_TREND = BASE_CTE + f"""
, actuals_5 AS (
    SELECT quarter_5 AS quarter,
        countDistinct(deal_id) AS actual_5,
        sum(amount)            AS amount_5
    FROM pipe_gen
    WHERE {_5PCT_FILTER}
      AND quarter_5 IS NOT NULL
    GROUP BY quarter_5
),
actuals_10 AS (
    SELECT quarter_10 AS quarter,
        countDistinct(deal_id) AS actual_10
    FROM pipe_gen
    WHERE {_10PCT_FILTER}
      AND quarter_10 IS NOT NULL
    GROUP BY quarter_10
),
actuals_20 AS (
    SELECT qual_quarter AS quarter,
        countDistinct(deal_id) AS actual_20
    FROM pipe_gen
    WHERE {_20PCT_FILTER}
      AND qual_quarter IS NOT NULL
    GROUP BY qual_quarter
),

tgts AS (
    SELECT 
        quarter,

        SUM(toFloat32(deals_target_5_l1))           AS l1_5,
        SUM(toFloat32(deals_target_5))              AS l2_5,
        SUM(toFloat32(deals_target_5_committed))    AS com_5,

        SUM(toFloat32(deals_target_10_l1))          AS l1_10,
        SUM(toFloat32(deals_target_10))             AS l2_10,
        SUM(toFloat32(deals_target_10_committed))   AS com_10,

        SUM(toFloat32(deals_target_20_l1))          AS l1_20,
        SUM(toFloat32(deals_target_20))             AS l2_20,
        SUM(toFloat32(deals_target_20_committed))   AS com_20

    FROM kore_ai_hubspot.gs_pipeline_quotas_v1
    WHERE {{TARGET_FILTERS}}
    GROUP BY quarter
)

SELECT
    t.quarter                                                              AS quarter,
    coalesce(a5.actual_5,   0) AS actual_deals,
    coalesce(a5.amount_5,   0) AS actual_amount,
    coalesce(t.l1_5,        0) AS l1_target_deals,
    coalesce(t.l2_5,        0) AS l2_target_deals,
    coalesce(t.com_5,       0) AS committed_target_deals,
    round(100.0 * coalesce(a5.actual_5, 0) / nullIf(t.l1_5, 0), 1)      AS pct_l1,
    round(100.0 * coalesce(a5.actual_5, 0) / nullIf(t.l2_5, 0), 1)      AS pct_l2,
    coalesce(a10.actual_10, 0) AS actual_10_deals,
    coalesce(t.l1_10,       0) AS l1_10_target,
    round(100.0 * coalesce(a10.actual_10,0) / nullIf(t.l1_10, 0), 1)    AS pct_l1_10,
    coalesce(a20.actual_20, 0) AS actual_20_deals,
    coalesce(t.l1_20,       0) AS l1_20_target,
    round(100.0 * coalesce(a20.actual_20,0) / nullIf(t.l1_20, 0), 1)    AS pct_l1_20
FROM tgts t
LEFT JOIN actuals_5  a5  ON a5.quarter  = t.quarter
LEFT JOIN actuals_10 a10 ON a10.quarter = t.quarter
LEFT JOIN actuals_20 a20 ON a20.quarter = t.quarter
WHERE t.quarter IS NOT NULL
ORDER BY CASE t.quarter
    WHEN 'Q1' THEN 1 WHEN 'Q2' THEN 2
    WHEN 'Q3' THEN 3 WHEN 'Q4' THEN 4 END
"""

# =============================================================================
# QUERY 3 — SLIDES 4 & 5: Region + Source Performance
# =============================================================================
QUERY_REGION_SOURCE_PERFORMANCE = BASE_CTE + f"""
, actuals_5 AS (
    SELECT region, deal_source_rollup,
        countDistinct(deal_id) AS deals_5,
        sum(amount)            AS amount_5
    FROM pipe_gen
    WHERE {_5PCT_FILTER}
    GROUP BY region, deal_source_rollup
),
actuals_10 AS (
    SELECT region, deal_source_rollup,
        countDistinct(deal_id) AS deals_10,
        sum(amount)            AS amount_10
    FROM pipe_gen
    WHERE {_10PCT_FILTER}
    GROUP BY region, deal_source_rollup
),
actuals_20 AS (
    SELECT region, deal_source_rollup,
        countDistinct(deal_id) AS deals_20,
        sum(amount)            AS amount_20
    FROM pipe_gen
    WHERE {_20PCT_FILTER}
    GROUP BY region, deal_source_rollup
),
tgts AS (
    SELECT
        region,
        CASE
            WHEN source IN ('Hyperscaler', 'Hyperscalers')
                                                                THEN 'Hyperscaler'
            WHEN source IN ('Executive Outreach', 'Investor')
                                                                THEN 'Executive Outreach'
            WHEN source IN ('Partner', 'Partner - Excluding Hyperscalers', 'Partner - Non Hyperscaler')
                                                                THEN 'Partner - Non Hyperscaler'
            WHEN source IN ('BDR Outbound', 'BDR')
                                                                THEN 'BDR'
            WHEN source IN ('Marketing', 'Customer Success', 'Inception', 'AE Outbound')
                                                                THEN source
            ELSE 'Other'
        END                                                      AS deal_source_rollup,
        SUM(toFloat32(deals_target_5_l1))  AS l1_5,
        SUM(toFloat32(deals_target_10_l1)) AS l1_10,
        SUM(toFloat32(deals_target_20_l1)) AS l1_20
    FROM kore_ai_hubspot.gs_pipeline_quotas_v1
    WHERE {{TARGET_FILTERS}}
    GROUP BY region, deal_source_rollup
),
all_combos AS (
    SELECT region, deal_source_rollup FROM actuals_5
    UNION DISTINCT
    SELECT region, deal_source_rollup FROM actuals_10
    UNION DISTINCT
    SELECT region, deal_source_rollup FROM actuals_20
    UNION DISTINCT
    SELECT region, deal_source_rollup FROM tgts
)
SELECT
    ac.region                                                         AS region,
    ac.deal_source_rollup                                             AS deal_source_rollup,
    coalesce(a5.deals_5,   0) AS deals_5,
    coalesce(a5.amount_5,  0) AS amount_5,
    coalesce(t.l1_5,       0) AS l1_5,
    round(100.0 * coalesce(a5.deals_5,  0) / nullIf(t.l1_5,  0), 1) AS pct_l1_5,
    coalesce(a10.deals_10, 0) AS deals_10,
    coalesce(a10.amount_10,0) AS amount_10,
    coalesce(t.l1_10,      0) AS l1_10,
    round(100.0 * coalesce(a10.deals_10,0) / nullIf(t.l1_10, 0), 1) AS pct_l1_10,
    coalesce(a20.deals_20, 0) AS deals_20,
    coalesce(a20.amount_20,0) AS amount_20,
    coalesce(t.l1_20,      0) AS l1_20,
    round(100.0 * coalesce(a20.deals_20,0) / nullIf(t.l1_20, 0), 1) AS pct_l1_20
FROM all_combos ac
LEFT JOIN actuals_5  a5  ON ac.region = a5.region  AND ac.deal_source_rollup = a5.deal_source_rollup
LEFT JOIN actuals_10 a10 ON ac.region = a10.region AND ac.deal_source_rollup = a10.deal_source_rollup
LEFT JOIN actuals_20 a20 ON ac.region = a20.region AND ac.deal_source_rollup = a20.deal_source_rollup
LEFT JOIN tgts       t   ON ac.region = t.region   AND ac.deal_source_rollup = t.deal_source_rollup
ORDER BY coalesce(a5.deals_5, 0) DESC, ac.region, ac.deal_source_rollup
"""

# =============================================================================
# QUERY 4 — SLIDE 6: Stage Velocity + Health Counts
# Uses 1% cohort scope (all FY2026 deals by create_date)
# =============================================================================
QUERY_STAGE_VELOCITY = BASE_CTE + f"""
SELECT
    deal_stage,
    avg_days_benchmark,
    round(avg(days_in_current_stage), 0)             AS avg_days_actual,
    countDistinct(deal_id)                           AS total_deals,
    countDistinctIf(deal_id, deal_health = 'Green')  AS green_deals,
    countDistinctIf(deal_id, deal_health = 'Yellow') AS yellow_deals,
    countDistinctIf(deal_id, deal_health = 'Red')    AS red_deals
FROM pipe_gen
WHERE {_1PCT_FILTER}
  AND deal_stage NOT IN (
        'Closed Won', 'Closed Lost', 'Prospect Disengaged',
        'Didn''t Qualify', 'Deal on Hold', '90% - Deal Desk Review'
      )
  AND days_in_current_stage IS NOT NULL
GROUP BY deal_stage, avg_days_benchmark
ORDER BY CASE deal_stage
    WHEN '1% - IQM Scheduled'      THEN 1
    WHEN '5% - IQM Held'           THEN 2
    WHEN '10% - Discovery'         THEN 3
    WHEN '20% - Solution'          THEN 4
    WHEN '30% - Proof'             THEN 5
    WHEN '40% - Proposal'          THEN 6
    WHEN '60% - Price Negotiation' THEN 7
    WHEN '75% - Contract Review'   THEN 8 END
"""

# =============================================================================
# QUERY 5 — SLIDE 7: Industry Mix + AI for X
# =============================================================================
QUERY_INDUSTRY_PRODUCT = BASE_CTE + f"""
SELECT
    kore_primary_industry,
    ai_for_x,
    countDistinctIf(deal_id, {_10PCT_FILTER}) AS deals_10pct,
    countDistinctIf(deal_id, {_20PCT_FILTER}) AS deals_20pct,
    sumIf(amount,            {_20PCT_FILTER}) AS amount_20pct
FROM pipe_gen
GROUP BY kore_primary_industry, ai_for_x
ORDER BY deals_20pct DESC
"""

# =============================================================================
# QUERY 6 — SLIDE 8: MTD / QTD / YTD Attainment
# Each cohort uses its own quarter column for QTD slicing
# =============================================================================
QUERY_PERIOD_ATTAINMENT = BASE_CTE + f"""
, actuals AS (
    SELECT
        -- 5% MTD/QTD/YTD
        countDistinctIf(deal_id, {_5PCT_FILTER}
            AND toYYYYMM(became_5_deal_date) = toYYYYMM(CURRENT_DATE()))         AS mtd_5,
        countDistinctIf(deal_id, {_5PCT_FILTER}
            AND quarter_5 = {_CURRENT_QUARTER})                                  AS qtd_5,
        countDistinctIf(deal_id, {_5PCT_FILTER})                                 AS ytd_5,
        sumIf(amount,            {_5PCT_FILTER})                                 AS ytd_5_amount,
        -- 10% MTD/QTD/YTD
        countDistinctIf(deal_id, {_10PCT_FILTER}
            AND toYYYYMM(became_10_deal_date) = toYYYYMM(CURRENT_DATE()))        AS mtd_10,
        countDistinctIf(deal_id, {_10PCT_FILTER}
            AND quarter_10 = {_CURRENT_QUARTER})                                 AS qtd_10,
        countDistinctIf(deal_id, {_10PCT_FILTER})                                AS ytd_10,
        sumIf(amount,            {_10PCT_FILTER})                                AS ytd_10_amount,
        -- 20% MTD/QTD/YTD
        countDistinctIf(deal_id, {_20PCT_FILTER}
            AND toYYYYMM(became_20_deal_date) = toYYYYMM(CURRENT_DATE()))        AS mtd_20,
        countDistinctIf(deal_id, {_20PCT_FILTER}
            AND qual_quarter = {_CURRENT_QUARTER})                               AS qtd_20,
        countDistinctIf(deal_id, {_20PCT_FILTER})                                AS ytd_20,
        sumIf(amount,            {_20PCT_FILTER})                                AS ytd_20_amount
    FROM pipe_gen
),
tgts AS (
    SELECT
        SUM(toFloat32(deals_target_5_l1))         AS l1_ytd_5,
        SUM(toFloat32(deals_target_5))            AS l2_ytd_5,
        SUM(toFloat32(deals_target_5_committed))  AS com_ytd_5,
        SUM(toFloat32(deals_target_10_l1))        AS l1_ytd_10,
        SUM(toFloat32(deals_target_10))           AS l2_ytd_10,
        SUM(toFloat32(deals_target_10_committed)) AS com_ytd_10,
        SUM(toFloat32(deals_target_20_l1))        AS l1_ytd_20,
        SUM(toFloat32(deals_target_20))           AS l2_ytd_20,
        SUM(toFloat32(deals_target_20_committed)) AS com_ytd_20,
        sumIf(toFloat32(deals_target_5_l1),
            month = {_CURRENT_MONTH})                                            AS l1_mtd_5,
        sumIf(toFloat32(deals_target_10_l1),
            month = {_CURRENT_MONTH})                                            AS l1_mtd_10,
        sumIf(toFloat32(deals_target_20_l1),
            month = {_CURRENT_MONTH})                                            AS l1_mtd_20,
        sumIf(toFloat32(deals_target_5_l1),
            quarter = {_CURRENT_QUARTER})                                        AS l1_qtd_5,
        sumIf(toFloat32(deals_target_10_l1),
            quarter = {_CURRENT_QUARTER})                                        AS l1_qtd_10,
        sumIf(toFloat32(deals_target_20_l1),
            quarter = {_CURRENT_QUARTER})                                        AS l1_qtd_20
    FROM kore_ai_hubspot.gs_pipeline_quotas_v1
    WHERE {{TARGET_FILTERS}}
)
SELECT
    a.mtd_5,  a.qtd_5,  a.ytd_5,  a.ytd_5_amount,
    a.mtd_10, a.qtd_10, a.ytd_10, a.ytd_10_amount,
    a.mtd_20, a.qtd_20, a.ytd_20, a.ytd_20_amount,
    t.l1_ytd_5,  t.l2_ytd_5,  t.com_ytd_5,
    t.l1_ytd_10, t.l2_ytd_10, t.com_ytd_10,
    t.l1_ytd_20, t.l2_ytd_20, t.com_ytd_20,
    t.l1_mtd_5,  t.l1_mtd_10, t.l1_mtd_20,
    t.l1_qtd_5,  t.l1_qtd_10, t.l1_qtd_20,
    round(100.0 * a.ytd_5  / nullIf(t.l1_ytd_5,  0), 1) AS pct_l1_ytd_5,
    round(100.0 * a.ytd_5  / nullIf(t.l2_ytd_5,  0), 1) AS pct_l2_ytd_5,
    round(100.0 * a.ytd_5  / nullIf(t.com_ytd_5, 0), 1) AS pct_com_ytd_5,
    round(100.0 * a.ytd_10 / nullIf(t.l1_ytd_10, 0), 1) AS pct_l1_ytd_10,
    round(100.0 * a.ytd_10 / nullIf(t.l2_ytd_10, 0), 1) AS pct_l2_ytd_10,
    round(100.0 * a.ytd_10 / nullIf(t.com_ytd_10,0), 1) AS pct_com_ytd_10,
    round(100.0 * a.ytd_20 / nullIf(t.l1_ytd_20, 0), 1) AS pct_l1_ytd_20,
    round(100.0 * a.ytd_20 / nullIf(t.l2_ytd_20, 0), 1) AS pct_l2_ytd_20,
    round(100.0 * a.ytd_20 / nullIf(t.com_ytd_20,0), 1) AS pct_com_ytd_20,
    round(100.0 * a.mtd_5  / nullIf(t.l1_mtd_5,  0), 1) AS pct_l1_mtd_5,
    round(100.0 * a.mtd_10 / nullIf(t.l1_mtd_10, 0), 1) AS pct_l1_mtd_10,
    round(100.0 * a.mtd_20 / nullIf(t.l1_mtd_20, 0), 1) AS pct_l1_mtd_20,
    round(100.0 * a.qtd_5  / nullIf(t.l1_qtd_5,  0), 1) AS pct_l1_qtd_5,
    round(100.0 * a.qtd_10 / nullIf(t.l1_qtd_10, 0), 1) AS pct_l1_qtd_10,
    round(100.0 * a.qtd_20 / nullIf(t.l1_qtd_20, 0), 1) AS pct_l1_qtd_20
FROM actuals a
CROSS JOIN tgts t
"""

# =============================================================================
# QUERY 7 — SLIDE 9: Deals to Watch
# =============================================================================
QUERY_DEALS_TO_WATCH = BASE_CTE + f"""
SELECT
    deal_id,
    deal_name,
    deal_owner_name  AS owner,
    team,
    deal_stage,
    region,
    amount,
    close_date,
    days_in_current_stage,
    avg_days_benchmark,
    deal_health,
    ai_for_x,
    kore_primary_industry,
    deal_source_rollup,
    deal_url,
    BANT
FROM pipe_gen
WHERE amount >= 1000000
  AND deal_health IN ('Red', 'Yellow')
  AND deal_stage NOT IN (
        'Closed Won', 'Closed Lost', 'Prospect Disengaged',
        'Didn''t Qualify', 'Deal on Hold', '90% - Deal Desk Review'
      )
  AND (
        ({_1PCT_FILTER})           -- 1%/pre-5% stage deals
        OR ({_5PCT_FILTER})        -- 5%+ stage deals
        OR ({_20PCT_ACTIVE_FILTER}) -- 20%+ active deals
      )
ORDER BY
    CASE deal_health WHEN 'Red' THEN 1 WHEN 'Yellow' THEN 2 END,
    amount DESC
LIMIT 20
"""

# =============================================================================
# QUERY 8 — Won & Lost Deals with Notes
# Won: top 5 Closed Won from 20% cohort (highest value, notes present)
# Lost: top 10 exited deals from 5% cohort (highest value, notes present)
# =============================================================================
QUERY_WON_LOST_DEALS = """
SELECT
    toInt64(d.deal_id) AS deal_id,
    d.deal_name,
    d.deal_stage,
    CASE
        WHEN d.region = 'india___sea' THEN 'ISEA'
        WHEN d.region = 'Africa'      THEN 'Middle East'
        WHEN d.region = 'japac'       THEN 'JAPAC'
        ELSE d.region
    END                                                                       AS region,
    CASE
        WHEN d.deal_source_rollup IN ('Executive Outreach', 'Investor')       THEN 'Executive Outreach'
        WHEN d.deal_source_rollup IN ('Marketing', 'Customer Success',
                                      'AE Outbound', 'Inception',
                                      'Hyperscaler')                          THEN d.deal_source_rollup
        WHEN d.deal_source_rollup IN ('BDR Outbound')                         THEN 'BDR'
        WHEN d.deal_source_rollup IN ('Partner')                              THEN 'Partner - Non Hyperscaler'
        ELSE 'Other'
    END                                                                       AS deal_source_rollup,
    CASE WHEN d.ai_for_x IS NULL THEN 'N/A' ELSE d.ai_for_x END              AS ai_for_x,
    CASE
        WHEN d.kore_primary_industry IN ('Financial Services',
                                         'Banking', 'Insurance')              THEN 'Financial Services'
        WHEN d.kore_primary_industry IN ('Manufacturing Discreet',
                                         'Manufacturing Process', 'CPG')      THEN 'Manufacturing'
        WHEN d.kore_primary_industry IN ('Hi-Tech',
                                         'Telecom / Media / Entertainment')   THEN 'TMT'
        WHEN d.kore_primary_industry IS NULL
          OR d.kore_primary_industry IN ('Business Services', 'Government',
                                         'Energy & Utilities', 'Education',
                                         'Restaurants', 'null', 'Energy')     THEN 'Other'
        ELSE d.kore_primary_industry
    END                                                                       AS kore_primary_industry,
    d.amount,
    CAST(LEFT(coalesce(d.close_date, '1900-01-01'), 10) AS DATE)              AS close_date,
    d.won_loss_notes,
    d.primary_closed_won_reason_,
    d.primary_closed_lost_reason,
    d.competitors,
    d.competition,
    CASE
        WHEN d.deal_stage IN ('Closed Won', '90% - Deal Desk Review') THEN 'won'
        WHEN d.deal_stage IN ('Closed Lost', 'Didn''t Qualify', 'Prospect Disengaged') THEN 'lost'
        ELSE NULL
    END                                                                       AS outcome
FROM hs_analytics.deals d FINAL
WHERE d.pipeline = 'default'
  AND d.deal_type NOT IN ('Partner-Led SMB')
  AND toInt64(d.deal_id) IN (
        SELECT DISTINCT toInt64(deal_id_hs)
        FROM kore_ai_hubspot.gs_deal_ids_hs
      )
"""


# =============================================================================
# MASTER FUNCTION
# =============================================================================

def get_all_pipeline_metrics(client):
    target_filters = "1=1"
    def run(sql, label):
        try:
            r = client.query(sql.replace("{TARGET_FILTERS}", target_filters))
            rows = [dict(zip(r.column_names, row)) for row in r.result_rows]
            print(f"✅ {label}: {len(rows)} row(s)")
            return rows
        except Exception as e:
            print(f"❌ ERROR in {label}: {e}")
            raise e

    return {
        "funnel":            (run(QUERY_FUNNEL_OVERVIEW,           "Slide 2 — Funnel") or [{}])[0],
        "quarterly_trend":    run(QUERY_QUARTERLY_TREND,           "Slide 3 — Quarterly Trend"),
        "region_source":      run(QUERY_REGION_SOURCE_PERFORMANCE, "Slides 4&5 — Region/Source"),
        "stage_velocity":     run(QUERY_STAGE_VELOCITY,            "Slide 6 — Stage Velocity"),
        "industry_product":   run(QUERY_INDUSTRY_PRODUCT,          "Slide 7 — Industry/Product"),
        "period_attainment": (run(QUERY_PERIOD_ATTAINMENT,         "Slide 8 — MTD/QTD/YTD") or [{}])[0],
        "deals_to_watch":     run(QUERY_DEALS_TO_WATCH,            "Slide 9 — Deals to Watch"),
        "funnel_conversions": (run(QUERY_FUNNEL_CONVERSIONS, "Funnel Conversions") or [{}])[0],
        "won_lost_deals": run(QUERY_WON_LOST_DEALS.replace("{TARGET_FILTERS}", target_filters), "Won/Lost Deals")
    }
