import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- COLOR PALETTE & STYLES ---
NAVY_HEADER = "1B365D"
DARK_NAVY = "0F2537"
ACCENT_BLUE = "2B6CB0"
ZEBRA_FILL = "F8FAFC"
WHITE = "FFFFFF"
BORDER_COLOR = "CBD5E1"

SEVERITY_STYLES = {
    "Critical": {
        "fill": PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="991B1B")
    },
    "High": {
        "fill": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="92400E")
    },
    "Medium": {
        "fill": PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="075985")
    },
    "Low": {
        "fill": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="166534")
    }
}

thin_border = Border(
    left=Side(style='thin', color=BORDER_COLOR),
    right=Side(style='thin', color=BORDER_COLOR),
    top=Side(style='thin', color=BORDER_COLOR),
    bottom=Side(style='thin', color=BORDER_COLOR)
)

header_border = Border(
    left=Side(style='thin', color="475569"),
    right=Side(style='thin', color="475569"),
    top=Side(style='medium', color="0F172A"),
    bottom=Side(style='medium', color="0F172A")
)

# --- DATASETS ---

DATA_AMBIGUOUS = [
    {
        "id": "AMB-01",
        "category": "Ambiguous Terms",
        "excerpt": "Prompt the model to draft Given-When-Then acceptance criteria...",
        "finding": "Undefined Model Specification: 'The model' does not specify provider, model family (e.g. GPT-4o, Claude 3.5 Sonnet, Gemini 1.5 Pro, Llama 3), version/checkpoint, temperature, top_p, or context limit. Different models and temperatures produce drastically different criteria.",
        "impact": "High test instability, non-reproducible outputs, and divergent acceptance criteria generated across test environments.",
        "severity": "High",
        "po_question": "Which specific AI model, version/release tag, temperature, and inference parameters are mandated, and is prompting triggered via UI or automated API?",
        "recommendation": "Explicitly define model parameters (e.g. Model: gpt-4o-2024-08-06, Temperature: 0.0, Seed: 42) in the configuration specification."
    },
    {
        "id": "AMB-02",
        "category": "Ambiguous Terms",
        "excerpt": "...based on its own findings",
        "finding": "Undefined Source of 'Findings': What constitutes a 'finding'? Is it raw text from a prior discovery chat, static code analysis output, bug logs, customer feedback transcripts, or unverified hallucinations? There is no input schema or validation.",
        "impact": "Garbage-in, garbage-out. The model may generate criteria grounded in hallucinated findings from an earlier turn.",
        "severity": "Critical",
        "po_question": "What is the exact data source, schema, and storage format for 'its own findings', and what upstream validation must occur before they feed into criteria generation?",
        "recommendation": "Specify a strict JSON schema contract for input findings with mandatory fields (Finding_ID, Category, Source, Verification_Status)."
    },
    {
        "id": "AMB-03",
        "category": "Ambiguous Terms",
        "excerpt": "Now inspect carefully: how many specific values did the model invent?",
        "finding": "Ambiguous Actor and Process in 'Inspect': Unclear whether 'inspect' is a manual human review (by QA, BA, or PO) or an automated parsing/regex/entity-extraction evaluation step. 'Carefully' has no operational definition.",
        "impact": "Test design cannot determine whether to build an automated NLP validation assertion or a human approval workflow screen.",
        "severity": "High",
        "po_question": "Who or what performs the inspection—is this an automated evaluation script/judge model or a manual human-in-the-loop review step in the UI?",
        "recommendation": "Clarify the workflow step: define automated NER-based value extraction vs human sign-off checkpoint."
    },
    {
        "id": "AMB-04",
        "category": "Ambiguous Terms",
        "excerpt": "...how many specific values did the model invent? ... Every invented value...",
        "finding": "Ambiguous Definition of 'Invented Value': Does 'invented value' refer exclusively to quantitative scalars (numbers, durations, percentages, currencies) or also unmentioned domain entities, roles (e.g., 'SuperAdmin'), statuses, and external integrations?",
        "impact": "False positives or false negatives during test evaluation depending on how 'invention' is classified.",
        "severity": "High",
        "po_question": "What is the formal definition of an 'invented value'—is it strictly numerical/temporal constants (e.g. 14 days, $50), or does it include novel business terms, roles, and status states?",
        "recommendation": "Define an explicit taxonomy: Level 1 (Quantities/Durations), Level 2 (Domain Entities/Roles), Level 3 (System Behaviors/Actions)."
    },
    {
        "id": "AMB-05",
        "category": "Ambiguous Terms",
        "excerpt": "...a business decision the model silently made on behalf of your organisation",
        "finding": "Ambiguity of 'Silently Made': What constitutes 'silent'? If the model outputs reasoning/Chain of Thought stating 'I assumed a 14-day window based on standard practice', is that silent? Or does silent mean any value not in the input prompt?",
        "impact": "Cannot write deterministic verification rules without knowing whether model explanations count as non-silent disclosures.",
        "severity": "Medium",
        "po_question": "How do we distinguish between a 'silent' assumption and an explicitly flagged assumption or recommendation in the model output?",
        "recommendation": "Require the model to return structured output with separate 'acceptance_criteria' and 'inferred_assumptions_log' fields."
    },
    {
        "id": "AMB-06",
        "category": "Ambiguous Terms",
        "excerpt": "This is the single most important observation of today's session.",
        "finding": "Ambiguous Concept of 'Session': Does 'session' denote an interactive training/workshop session, a user login session token, or an automated test execution run? 'Most important' is subjective and untestable.",
        "impact": "Scope confusion: Is this a functional software requirement for a product feature, or an educational exercise/workshop prompt?",
        "severity": "Low",
        "po_question": "Does 'session' refer to an application runtime session requiring telemetry/reporting, or is this prompt meant for a training/demonstration context?",
        "recommendation": "State the operational context: If production software, define session ID tracking and alert logging requirements."
    }
]

DATA_MISSING_CONDITIONS = [
    {
        "id": "MC-01",
        "category": "Missing Conditions",
        "excerpt": "Prompt the model to draft Given-When-Then acceptance criteria...",
        "finding": "Model Failure, Timeout, and Error Handling: No branch exists for when the LLM service returns HTTP 429 (rate limit), HTTP 500/503 (server error), network timeout, or context window overflow.",
        "impact": "System crashes or hangs indefinitely; unhandled exception during automated testing or user execution.",
        "severity": "High",
        "po_question": "What is the expected system behavior, retry policy, timeout threshold (e.g. 30s), and user-facing error message when model generation fails?",
        "recommendation": "Define an explicit exception handling branch: retry up to 3 times with exponential backoff, then return a structured error response."
    },
    {
        "id": "MC-02",
        "category": "Missing Conditions",
        "excerpt": "...how many specific values did the model invent?",
        "finding": "Zero Invented Values Branch (Pass Condition): The requirement emphasizes what happens when values ARE invented, but fails to define the workflow branch when the count of invented values is 0.",
        "impact": "No clear completion state for the happy path where acceptance criteria are clean and fully grounded in findings.",
        "severity": "Medium",
        "po_question": "What is the expected automated action or workflow state when zero invented values are found (e.g. auto-approve to Jira, mark as Verified)?",
        "recommendation": "Add positive branch: 'If invented_values_count == 0, transition story state to Ready for Review and tag as Grounded.'"
    },
    {
        "id": "MC-03",
        "category": "Missing Conditions",
        "excerpt": "Every invented value is a business decision the model silently made on behalf of your organisation",
        "finding": "Gating and Remediation Threshold: What happens when invented values are detected? Is there a tolerance threshold (e.g., 1 warning vs hard block)? Can a human PO accept/adopt the invented value into the official business specification?",
        "impact": "No defect triage pathway; cannot determine if the system should reject the criteria, flag them for review, or prompt the user for confirmation.",
        "severity": "Critical",
        "po_question": "Does the detection of any invented value hard-block acceptance, and can a Product Owner override/adopt an invented value as a valid business rule?",
        "recommendation": "Define triage states: Flagged for Review -> PO Options: [1] Accept as New Business Rule, [2] Reject and Re-prompt with Constraint, [3] Edit Manually."
    },
    {
        "id": "MC-04",
        "category": "Missing Conditions",
        "excerpt": "...draft Given-When-Then acceptance criteria based on its own findings",
        "finding": "Malformed Output / Invalid Gherkin Syntax: No condition addresses model outputs that fail standard Gherkin syntax (e.g., missing 'Given', missing 'Then', malformed markdown, or code dumps).",
        "impact": "Automated BDD parsers (Cucumber, Behave, SpecFlow) will crash upon ingesting malformed output.",
        "severity": "High",
        "po_question": "What is the validation criteria for Gherkin syntax compliance, and how should malformed outputs be rejected or auto-repaired?",
        "recommendation": "Implement an automated Gherkin linter validation gate before acceptance criteria can be passed to downstream steps."
    },
    {
        "id": "MC-05",
        "category": "Missing Conditions",
        "excerpt": "Prompt the model to draft Given-When-Then acceptance criteria...",
        "finding": "Stochastic Non-Determinism Across Runs: If the prompt is re-run on the same findings, the model may invent different values (e.g., 14 days on Run 1, 30 days on Run 2, no duration on Run 3).",
        "impact": "Flaky test results; lack of audit reproducibility and fluctuating compliance status.",
        "severity": "High",
        "po_question": "How should non-determinism be managed—should the system enforce strict temperature=0 with fixed seeds, or store multi-run samples?",
        "recommendation": "Mandate temperature=0.0 and fixed random seeds for deterministic output, and log prompt-output pairings with SHA-256 hashes."
    },
    {
        "id": "MC-06",
        "category": "Missing Conditions",
        "excerpt": "...how many specific values did the model invent? ... Every invented value...",
        "finding": "Persistence and Audit Trail Schema: The requirement states to observe and count invented values, but provides no requirement for where this observation is saved, logged, or reported.",
        "impact": "Loss of audit compliance data; unable to verify that inspections occurred or track AI governance metrics over time.",
        "severity": "Medium",
        "po_question": "Where should the log of detected invented values, model prompts, and review dispositions be persisted, and what audit retention period applies?",
        "recommendation": "Create an AI Audit Log table: [Execution_ID, Prompt_Hash, Model_Name, Generated_GWT, Invented_Values_JSON, Status, Reviewer_ID]."
    }
]

DATA_UNSTATED_ASSUMPTIONS = [
    {
        "id": "UBA-01",
        "category": "Unstated Assumptions",
        "excerpt": "...how many specific values did the model invent?",
        "finding": "Assumption of an Objective Ground Truth Document: Assumes an infallible baseline exists to prove whether a value was 'invented' or already documented. If upstream findings are fuzzy, determining 'invention' is subjective.",
        "impact": "QA testers cannot reliably judge if a value is an 'invention' or a legitimate synonym/derivation without an authoritative source.",
        "severity": "Critical",
        "po_question": "What is the baseline source of truth against which values are compared, and what document holds final authority in disputes?",
        "recommendation": "Establish a clear requirement hierarchy: PRD > Technical Spec > Raw Findings > AI Output."
    },
    {
        "id": "UBA-02",
        "category": "Unstated Assumptions",
        "excerpt": "Every invented value is a business decision the model silently made on behalf of your organisation",
        "finding": "Assumption that ALL Model Deductions are Negative/Defective: Assumes that any value not explicitly given in the input is an unacceptable hallucination. In many workflows, AI is expected to suggest standard best practices.",
        "impact": "Over-restrictive verification could reject valuable suggestions, or QA might flag intentional model recommendations as defects.",
        "severity": "High",
        "po_question": "Does the organization prohibit ALL unprompted value suggestions, or is there a distinction between 'Unauthorized Inventions' vs 'Proposed Defaults'?",
        "recommendation": "Formulate an AI Governance Policy categorizing model inferences into: Prohibited Hallucinations vs Flagged Recommendations for PO Approval."
    },
    {
        "id": "UBA-03",
        "category": "Unstated Assumptions",
        "excerpt": "A 14-day window? A 7-day refund period?",
        "finding": "Domain-Specific Bias Assumption: Assumes requirements are commercial/e-commerce oriented. If the system is an infrastructure, authentication, or medical system, invented parameters (e.g. 256-bit encryption, 30s session timeout) have different implications.",
        "impact": "Test rules tuned only for durations/refunds will miss critical technical or security hallucinations.",
        "severity": "Medium",
        "po_question": "Is this inspection framework limited to commercial/business rules, or must it encompass technical, security, and non-functional specifications?",
        "recommendation": "Generalize the inspection criteria to cover: [1] Business/Policy values, [2] Performance/SLA values, [3] Security/Access parameters."
    },
    {
        "id": "UBA-04",
        "category": "Unstated Assumptions",
        "excerpt": "...silently made on behalf of your organisation",
        "finding": "Assumption of Legal / RACI Responsibility: Assumes the model's output could bind the organization if uninspected, implying an unstated assumption about human sign-off accountability.",
        "impact": "Lack of defined human gatekeepers could lead to unreviewed acceptance criteria directly driving sprint delivery and release testing.",
        "severity": "High",
        "po_question": "What is the governance rule and RACI matrix for AI-generated acceptance criteria before they become legally and contractually binding for delivery?",
        "recommendation": "Establish mandatory human-in-the-loop sign-off (PO/QA Lead) before any AI-generated criteria can be transitioned to active sprints."
    }
]

DATA_VERIFIABILITY = [
    {
        "id": "NMV-01",
        "category": "Measurable Verifiability",
        "excerpt": "Now inspect carefully: how many specific values did the model invent?",
        "finding": "'Carefully' is Subjective and Non-Verifiable: In QA test design, 'carefully' cannot be asserted. A test cannot verify if a human or machine inspected 'carefully'—there is no quantifiable threshold, checklist, or SLA.",
        "impact": "Test case pass/fail criteria become subjective opinions, leading to unresolvable QA-developer disputes.",
        "severity": "High",
        "po_question": "What objective, measurable metrics define a valid inspection (e.g. 100% token entity comparison, cross-reference against approved term glossary)?",
        "recommendation": "Replace 'carefully' with verifiable criteria: 'Inspect against the Source Entity Lexicon; achieve 0 unflagged literal mismatches.'"
    },
    {
        "id": "NMV-02",
        "category": "Measurable Verifiability",
        "excerpt": "This is the single most important observation of today's session.",
        "finding": "'Single Most Important Observation' is an Opinion: Software verification cannot assert 'importance' without a quantitative scoring system, severity scale, or telemetry flag.",
        "impact": "Cannot write an automated assertion for 'most important observation'; test will fail for ambiguity.",
        "severity": "Medium",
        "po_question": "How should this observation be formally classified in the application (e.g., Severity = Critical, Alert Priority = P1, Risk Score = 10)?",
        "recommendation": "Define a concrete telemetry field: `observation_severity = Enum('CRITICAL_BUSINESS_RISK')`."
    },
    {
        "id": "NMV-03",
        "category": "Measurable Verifiability",
        "excerpt": "...draft Given-When-Then acceptance criteria based on its own findings",
        "finding": "'Based On' Lacks Traceability Metrics: How does a QA engineer objectively verify that criterion X is 'based on' finding Y? Without explicit citation indices or semantic similarity thresholds, grounding cannot be verified.",
        "impact": "Subjective validation of whether acceptance criteria are faithful to findings or subtly deviating.",
        "severity": "High",
        "po_question": "What measurable metric or traceability mechanism proves that an acceptance criterion is directly grounded in the findings?",
        "recommendation": "Require bidirectional traceability: Each Given-When-Then scenario must include a `@source_finding_id` annotation."
    },
    {
        "id": "NMV-04",
        "category": "Measurable Verifiability",
        "excerpt": "...a business decision the model silently made on behalf of your organisation",
        "finding": "'Silently Made' Lacks Testable Telemetry: Without an event log specification, asserting that a decision was made 'silently' vs 'explicitly' is untestable in software automation.",
        "impact": "Automated regression tests cannot verify whether silent decisions were intercepted or logged.",
        "severity": "Medium",
        "po_question": "What system telemetry or audit record verifies whether a decision was made 'silently' versus accompanied by an explicit reasoning disclosure?",
        "recommendation": "Specify an event payload: `event_name: 'AI_UNGROUNDED_INFERENCE_DETECTED'`, with fields for entity, inferred value, and confidence."
    }
]

ALL_FINDINGS = DATA_AMBIGUOUS + DATA_MISSING_CONDITIONS + DATA_UNSTATED_ASSUMPTIONS + DATA_VERIFIABILITY

def create_styled_sheet(ws, title, subtitle, data, category_name=None):
    ws.views.sheetView[0].showGridLines = True
    
    # 1. Main Title Banner
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = f"QA REQUIREMENTS REVIEW — {title.upper()}"
    title_cell.font = Font(name="Calibri", size=15, bold=True, color=WHITE)
    title_cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # 2. Subtitle / Metadata
    ws.merge_cells("A2:H2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Role: Senior QA Engineer | Target: AI Acceptance Criteria Engine | Status: Gate 0 Pre-Test Design | {subtitle}"
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="334155")
    sub_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # 3. KPI Block
    crit_count = sum(1 for d in data if d["severity"] == "Critical")
    high_count = sum(1 for d in data if d["severity"] == "High")
    med_count = sum(1 for d in data if d["severity"] == "Medium")
    low_count = sum(1 for d in data if d["severity"] == "Low")
    total_count = len(data)

    kpis = [
        ("Total Items", total_count, "2563EB", "DBEAFE"),
        ("Critical Defects", crit_count, "DC2626", "FEE2E2"),
        ("High Risks", high_count, "D97706", "FEF3C7"),
        ("Medium Questions", med_count, "0284C7", "E0F2FE"),
        ("Low Notes", low_count, "16A34A", "DCFCE7")
    ]

    col_targets = ["B", "C", "D", "E", "F"]
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 26

    for (k_title, k_val, text_col, fill_col), col in zip(kpis, col_targets):
        lbl = ws[f"{col}4"]
        lbl.value = k_title.upper()
        lbl.font = Font(name="Calibri", size=8, bold=True, color="64748B")
        lbl.alignment = Alignment(horizontal="center", vertical="center")
        lbl.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        lbl.border = thin_border

        val = ws[f"{col}5"]
        val.value = k_val
        val.font = Font(name="Calibri", size=16, bold=True, color=text_col)
        val.alignment = Alignment(horizontal="center", vertical="center")
        val.fill = PatternFill(start_color=fill_col, end_color=fill_col, fill_type="solid")
        val.border = thin_border

    # 4. Table Headers
    headers = [
        ("Item ID", 10),
        ("Analysis Category", 18),
        ("Requirement Excerpt Under Review", 32),
        ("QA Analysis & Defect Description", 46),
        ("Defect Potential & Testing Risk", 32),
        ("Severity", 12),
        ("Specific Question for Product Owner", 48),
        ("Recommended Action / Spec Amendment", 38)
    ]

    header_row = 7
    ws.row_dimensions[header_row].height = 28

    for c_idx, (h_name, _) in enumerate(headers, start=1):
        c_letter = get_column_letter(c_idx)
        cell = ws[f"{c_letter}{header_row}"]
        cell.value = h_name
        cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center" if c_idx in [1, 2, 6] else "left", vertical="center", wrap_text=True)
        cell.border = header_border

    # 5. Populate Data Rows
    current_row = header_row + 1
    for row_idx, item in enumerate(data):
        ws.row_dimensions[current_row].height = 68
        bg_fill = PatternFill(start_color=WHITE if row_idx % 2 == 0 else ZEBRA_FILL,
                              end_color=WHITE if row_idx % 2 == 0 else ZEBRA_FILL,
                              fill_type="solid")

        row_data = [
            (item["id"], "center", Font(name="Calibri", size=10, bold=True)),
            (item["category"], "center", Font(name="Calibri", size=10)),
            (f'"{item["excerpt"]}"', "left", Font(name="Calibri", size=10, italic=True)),
            (item["finding"], "left", Font(name="Calibri", size=10)),
            (item["impact"], "left", Font(name="Calibri", size=10)),
            (item["severity"], "center", None),
            (item["po_question"], "left", Font(name="Calibri", size=10, bold=True, color="1E3A8A")),
            (item["recommendation"], "left", Font(name="Calibri", size=10, color="0F172A"))
        ]

        for c_idx, (val, align_h, font_style) in enumerate(row_data, start=1):
            c_letter = get_column_letter(c_idx)
            cell = ws[f"{c_letter}{current_row}"]
            cell.value = val
            cell.alignment = Alignment(horizontal=align_h, vertical="top", wrap_text=True)
            cell.border = thin_border

            if c_idx == 6:  # Severity
                cell.fill = SEVERITY_STYLES[val]["fill"]
                cell.font = SEVERITY_STYLES[val]["font"]
            else:
                cell.fill = bg_fill
                cell.font = font_style

        current_row += 1

    # Column Widths
    for c_idx, (_, col_width) in enumerate(headers, start=1):
        c_letter = get_column_letter(c_idx)
        ws.column_dimensions[c_letter].width = col_width

    # Freeze Panes
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:H{current_row - 1}"


def create_executive_summary_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "QA REQUIREMENT ANALYSIS & GATE 0 REVIEW REPORT"
    t.font = Font(name="Calibri", size=16, bold=True, color=WHITE)
    t.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42

    # Executive Summary Card
    ws.merge_cells("A3:G3")
    sec1 = ws["A3"]
    sec1.value = "1. EXECUTIVE QA ASSESSMENT & GATE DECISION"
    sec1.font = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)

    ws.merge_cells("A4:G5")
    box = ws["A4"]
    box.value = (
        "GATE 0 DECISION: REJECT / HOLD BEFORE TEST DESIGN.\n"
        "The requirement under review establishes an AI prompting workflow for drafting Given-When-Then acceptance criteria, "
        "highlighting that 'Every invented value is a business decision the model silently made on behalf of your organisation.' "
        "However, as a functional specification, it lacks foundational architectural definitions, error branches, verification metrics, "
        "and governance controls. Writing test cases against this specification in its current state will lead to invalid tests and ungrounded automation."
    )
    box.font = Font(name="Calibri", size=10, bold=False, color="1E293B")
    box.fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
    box.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    box.border = thin_border

    # Core QA Findings Matrix (Row 7)
    ws.merge_cells("A7:G7")
    sec2 = ws["A7"]
    sec2.value = "2. DEFECT SUMMARY MATRIX ACROSS 4 CORE QA PILLARS"
    sec2.font = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)

    # Matrix Headers
    matrix_headers = ["Pillar / Dimension", "Total Findings", "Critical", "High", "Medium", "Low", "Primary QA Risk"]
    m_row = 8
    ws.row_dimensions[m_row].height = 24
    for c_idx, h in enumerate(matrix_headers, start=1):
        cell = ws[f"{get_column_letter(c_idx)}{m_row}"]
        cell.value = h
        cell.font = Font(name="Calibri", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        cell.alignment = Alignment(horizontal="center" if c_idx in [2, 3, 4, 5, 6] else "left", vertical="center")
        cell.border = thin_border

    matrix_data = [
        ("1. Ambiguous Terms", len(DATA_AMBIGUOUS), 1, 3, 1, 1, "Undefined model specs, fuzzy 'findings', subjective inspection"),
        ("2. Missing Conditions / Branches", len(DATA_MISSING_CONDITIONS), 1, 3, 2, 0, "No API failure handling, missing zero-invention pass branch, no Gherkin linter"),
        ("3. Unstated Business Assumptions", len(DATA_UNSTATED_ASSUMPTIONS), 1, 2, 1, 0, "No ground truth baseline, unclear whether defaults are permitted, unstated RACI"),
        ("4. Measurable Verifiability Issues", len(DATA_VERIFIABILITY), 0, 2, 2, 0, "'Carefully', 'most important', and 'silently' cannot be asserted objectively"),
        ("TOTALS ACROSS ALL PILLARS", len(ALL_FINDINGS), 3, 10, 6, 1, "Immediate PO clarification required before test case authoring begins")
    ]

    for r_idx, row_items in enumerate(matrix_data, start=9):
        ws.row_dimensions[r_idx].height = 24
        is_total = (r_idx == 13)
        for c_idx, val in enumerate(row_items, start=1):
            cell = ws[f"{get_column_letter(c_idx)}{r_idx}"]
            cell.value = val
            cell.border = thin_border
            if is_total:
                cell.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
                cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            else:
                cell.font = Font(name="Calibri", size=10, bold=(c_idx == 1))
                cell.fill = PatternFill(start_color=WHITE if r_idx % 2 == 1 else ZEBRA_FILL,
                                      end_color=WHITE if r_idx % 2 == 1 else ZEBRA_FILL,
                                      fill_type="solid")
            cell.alignment = Alignment(horizontal="center" if c_idx in [2, 3, 4, 5, 6] else "left", vertical="center")

    # Core QA Recommendation (Row 15)
    ws.merge_cells("A15:G15")
    sec3 = ws["A15"]
    sec3.value = "3. CRITICAL QA TAKEAWAY FOR THE PRODUCT OWNER & TEAM"
    sec3.font = Font(name="Calibri", size=12, bold=True, color=NAVY_HEADER)

    ws.merge_cells("A16:G18")
    takeaway = ws["A16"]
    takeaway.value = (
        "• THE OBSERVATION IS VALID BUT INCOMPLETE AS A SPECIFICATION:\n"
        "  The observation that 'Every invented value is a business decision the model silently made on behalf of your organisation' "
        "is a profound truth of Generative AI. It exposes the acute risk of unauthorized hallucinated business rules.\n"
        "• HOWEVER, FROM A QA PERSPECTIVE:\n"
        "  An observation is not an acceptance criterion. To make this requirement testable, the Product Owner must specify: "
        "(1) The exact schema and source of truth, (2) The detection mechanism (human vs automated judge), (3) The triage workflow when values are invented, "
        "and (4) The governance policy governing model inferences."
    )
    takeaway.font = Font(name="Calibri", size=9, bold=False, color="0F172A")
    takeaway.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    takeaway.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    takeaway.border = thin_border

    # Set column widths
    col_widths = [32, 16, 12, 12, 12, 12, 45]
    for idx, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w


def generate_all_files():
    base_path = r"c:\Users\Acer\Downloads\Antigravity"

    print("Generating Master Workbook...")
    wb_master = openpyxl.Workbook()
    
    # 1. Executive Summary Sheet
    ws_exec = wb_master.active
    ws_exec.title = "Executive Summary"
    create_executive_summary_sheet(ws_exec)

    # 2. Ambiguous Terms Sheet
    ws_amb = wb_master.create_sheet(title="1. Ambiguous Terms")
    create_styled_sheet(ws_amb, "1. Ambiguous Terms Analysis", "Identification of double-meanings and undefined entities", DATA_AMBIGUOUS)

    # 3. Missing Conditions Sheet
    ws_mc = wb_master.create_sheet(title="2. Missing Conditions")
    create_styled_sheet(ws_mc, "2. Missing Conditions & Branches", "Unhandled system errors, null paths, and gating branches", DATA_MISSING_CONDITIONS)

    # 4. Unstated Assumptions Sheet
    ws_uba = wb_master.create_sheet(title="3. Unstated Assumptions")
    create_styled_sheet(ws_uba, "3. Unstated Business-Rule Assumptions", "Implicit ground truth, domain biases, and governance assumptions", DATA_UNSTATED_ASSUMPTIONS)

    # 5. Verifiability Sheet
    ws_ver = wb_master.create_sheet(title="4. Verifiability Issues")
    create_styled_sheet(ws_ver, "4. Measurable Verifiability Review", "Subjective adverbs, unverifiable claims, and missing test assertions", DATA_VERIFIABILITY)

    # 6. PO Question & Triage Log Sheet
    ws_po = wb_master.create_sheet(title="5. PO Question Log")
    create_styled_sheet(ws_po, "Master Product Owner Question Log", "Consolidated 20-item review questions prioritized for triage", ALL_FINDINGS)

    master_file = f"{base_path}\\Requirement_Review_Master_Workbook.xlsx"
    wb_master.save(master_file)
    print(f"Master Workbook saved: {master_file}")

    # Generate Individual Workbooks for each category separately
    individual_configs = [
        ("1_Ambiguous_Terms_Review.xlsx", "Ambiguous Terms Review", "Category 1 of 4: Ambiguous and Multiple-Interpretation Terms", DATA_AMBIGUOUS),
        ("2_Missing_Conditions_Review.xlsx", "Missing Conditions Review", "Category 2 of 4: Unhandled Branches, Failure States, and Gating", DATA_MISSING_CONDITIONS),
        ("3_Unstated_Assumptions_Review.xlsx", "Unstated Assumptions Review", "Category 3 of 4: Implicit Business Rules, Ground Truth, and Governance", DATA_UNSTATED_ASSUMPTIONS),
        ("4_Measurable_Verifiability_Review.xlsx", "Measurable Verifiability Review", "Category 4 of 4: Subjective Statements & Untestable Elements", DATA_VERIFIABILITY),
        ("5_Product_Owner_Question_Log.xlsx", "Consolidated PO Question Log", "Complete 20-Item Question Matrix for Product Owner Sign-Off", ALL_FINDINGS)
    ]

    for fname, title, sub, data in individual_configs:
        wb_ind = openpyxl.Workbook()
        ws_ind = wb_ind.active
        ws_ind.title = title[:31]
        create_styled_sheet(ws_ind, title, sub, data)
        out_path = f"{base_path}\\{fname}"
        wb_ind.save(out_path)
        print(f"Saved: {out_path}")

if __name__ == "__main__":
    generate_all_files()
