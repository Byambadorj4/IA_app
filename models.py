# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin

db = SQLAlchemy()


# =========================================================
# USER
# =========================================================
class User(db.Model, UserMixin):
    __tablename__ = "user"

    id = db.Column(db.Integer, primary_key=True)

    username = db.Column(db.String(150), nullable=False, unique=True, index=True)
    email = db.Column(db.String(150), nullable=False, unique=True, index=True)

    password_hash = db.Column(db.String(255), nullable=False)

    is_admin = db.Column(db.Boolean, default=False)
    is_blocked = db.Column(db.Boolean, default=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self) -> str:
        return f"<User {self.id} {self.username}>"



# =========================================================
# ISSUE
# =========================================================
class Issue(db.Model):
    __tablename__ = "issue"

    id = db.Column(db.Integer, primary_key=True)

    company_name = db.Column(db.String(255))
    auditor_name = db.Column(db.String(255))

    issue = db.Column(db.Text)
    recommendation = db.Column(db.Text)

    detail_issue = db.Column(db.Text)
    issue_text = db.Column(db.Text)

    detail_context = db.Column(db.Text)
    detail_criteria = db.Column(db.Text)
    detail_cause = db.Column(db.Text)
    detail_impact = db.Column(db.Text)

    recommendation_number = db.Column(db.Text)
    report_name = db.Column(db.Text)

    reply_comment = db.Column(db.Text)

    implement_due_date = db.Column(db.Date)
    identified_date = db.Column(db.Date)
    materiality = db.Column(db.String(100))

    has_actual_loss = db.Column(db.String(50))
    money_amount = db.Column(db.String(100))

    evidence_file = db.Column(db.Text)

    risk_event_kind = db.Column(db.String(255))
    risk_factor = db.Column(db.String(255))
    risk_classification = db.Column(db.String(255))
    sub_risk_classification = db.Column(db.String(255))
    risk_owner = db.Column(db.String(255))

    controls_text = db.Column(db.Text)

    ctl_design = db.Column(db.String(50))
    ctl_operates = db.Column(db.String(50))
    ctl_awareness = db.Column(db.String(50))

    control_pct = db.Column(db.String(50))

    r5_financial = db.Column(db.String(50))
    r5_regulatory = db.Column(db.String(50))
    r5_stakeholder = db.Column(db.String(50))
    r5_operation = db.Column(db.String(50))
    r5_health = db.Column(db.String(50))
    r5_probability = db.Column(db.String(50))

    residual_score = db.Column(db.Integer)
    residual_level = db.Column(db.String(50))
    action_rank = db.Column(db.String(50))
    risk_category = db.Column(db.String(50))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Relationship
    followups = db.relationship(
        "FollowUp",
        back_populates="issue",
        cascade="all, delete-orphan",
        lazy=True
    )

    def __repr__(self) -> str:
        return f"<Issue {self.id} {self.company_name}>"



# =========================================================
# FOLLOW UP
# =========================================================
class FollowUp(db.Model):
    __tablename__ = "follow_up"

    id = db.Column(db.Integer, primary_key=True)

    issue_id = db.Column(
        db.Integer,
        db.ForeignKey("issue.id", ondelete="CASCADE"),
        nullable=False,
        index=True
    )

    action_type = db.Column(db.String(255))
    target_level = db.Column(db.String(50))

    action_plan = db.Column(db.Text)

    board_date = db.Column(db.Date)
    deadline = db.Column(db.Date)

    note = db.Column(db.Text)

    impl_percent = db.Column(db.String(50))
    target_hit = db.Column(db.String(50))

    auditor_conclusion = db.Column(db.Text)

    fully_implemented = db.Column(db.String(50))
    final_percent = db.Column(db.String(50))

    follow_up = db.Column(db.Text)

    auditor_name = db.Column(db.String(255))

    required_resource = db.Column(db.Text)
    required_budget = db.Column(db.Text)
    responsible_unit = db.Column(db.Text)

    on_time = db.Column(db.String(50))

    evidence_file = db.Column(db.Text)
    auditor_evidence_file = db.Column(db.Text)

    reduced_amount = db.Column(db.Text)

    residual_score = db.Column(db.Integer)
    residual_level = db.Column(db.String(50))
    action_rank = db.Column(db.String(50))

    r5_financial = db.Column(db.String(50))
    r5_regulatory = db.Column(db.String(50))
    r5_stakeholder = db.Column(db.String(50))
    r5_operation = db.Column(db.String(50))
    r5_health = db.Column(db.String(50))
    r5_probability = db.Column(db.String(50))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    issue = db.relationship(
        "Issue",
        back_populates="followups"
    )

    def __repr__(self) -> str:
        return f"<FollowUp {self.id} issue={self.issue_id}>"



# =========================================================
# LOG
# =========================================================
class Log(db.Model):
    __tablename__ = "log"

    id = db.Column(db.Integer, primary_key=True)

    user = db.Column(db.String(255))
    action = db.Column(db.Text)

    timestamp = db.Column(db.DateTime, default=datetime.utcnow)



# =========================================================
# GUIDELINE
# =========================================================
class Guideline(db.Model):
    __tablename__ = "guideline"

    id = db.Column(db.Integer, primary_key=True)

    company_name = db.Column(db.String(255))

    audit_type = db.Column(db.String(255))
    audit_subtype = db.Column(db.String(255))

    team_leader = db.Column(db.String(255))
    team_members = db.Column(db.Text)

    scope_start = db.Column(db.Date)
    scope_end = db.Column(db.Date)

    exec_start = db.Column(db.Date)
    exec_end = db.Column(db.Date)

    extended_end = db.Column(db.Date)
    extension_note = db.Column(db.Text)

    created_by = db.Column(db.String(255))

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    approved_pdf = db.Column(db.Text)

    def __repr__(self) -> str:
        return f"<Guideline {self.id} {self.company_name}>"