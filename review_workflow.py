"""
质量安全负责人审核模块 (PRD 5.12)
=================================
安全评估报告的审核工作流
"""

import json
import os
from datetime import datetime
from dataclasses import dataclass, field, asdict


REVIEWS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'userdata', 'reviews'
)


def _ensure_reviews_dir():
    os.makedirs(REVIEWS_DIR, exist_ok=True)


@dataclass
class ReviewAction:
    """审核操作记录"""
    action: str                           # submit / approve / reject / revise
    comment: str                          # 审核意见
    reviewer_name: str                    # 审核人
    timestamp: str = ''                   # 时间戳

    def __post_init__(self):
        if not self.timestamp:
            self.timestamp = datetime.now().isoformat()


@dataclass
class ReviewRecord:
    """审核记录"""
    review_id: str                        # 审核编号
    product_name: str                     # 产品名称
    report_number: str                    # 报告编号
    status: str = 'pending'               # pending / in_review / approved / rejected / revision_needed
    submitted_by: str = ''                # 提交人
    history: list = field(default_factory=list)  # list of ReviewAction
    notes: str = ''
    created_at: str = ''
    updated_at: str = ''

    def __post_init__(self):
        now = datetime.now().isoformat()
        if not self.created_at:
            self.created_at = now
        if not self.updated_at:
            self.updated_at = now

    def to_dict(self) -> dict:
        return {
            'review_id': self.review_id,
            'product_name': self.product_name,
            'report_number': self.report_number,
            'status': self.status,
            'submitted_by': self.submitted_by,
            'history': [asdict(a) for a in self.history],
            'notes': self.notes,
            'created_at': self.created_at,
            'updated_at': self.updated_at,
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ReviewRecord':
        record = cls(
            review_id=data['review_id'],
            product_name=data['product_name'],
            report_number=data['report_number'],
            status=data.get('status', 'pending'),
            submitted_by=data.get('submitted_by', ''),
            notes=data.get('notes', ''),
            created_at=data.get('created_at', ''),
            updated_at=data.get('updated_at', ''),
        )
        for h in data.get('history', []):
            record.history.append(ReviewAction(**h))
        return record


def create_review_id() -> str:
    """Generate a unique review ID."""
    now = datetime.now()
    return f'REV-{now.strftime("%Y%m%d")}-{now.strftime("%H%M%S")}'


def create_review(
    product_name: str,
    report_number: str,
    submitted_by: str = '',
) -> ReviewRecord:
    """Create a new review record."""
    return ReviewRecord(
        review_id=create_review_id(),
        product_name=product_name,
        report_number=report_number,
        submitted_by=submitted_by,
    )


def submit_review(record: ReviewRecord, reviewer_name: str = '') -> ReviewRecord:
    """Submit review for approval."""
    record.status = 'in_review'
    record.history.append(ReviewAction(
        action='submit',
        comment='提交安全评估报告审核',
        reviewer_name=reviewer_name,
    ))
    record.updated_at = datetime.now().isoformat()
    return record


def approve_review(record: ReviewRecord, reviewer_name: str, comment: str = '') -> ReviewRecord:
    """Approve the review."""
    record.status = 'approved'
    record.history.append(ReviewAction(
        action='approve',
        comment=comment or '审核通过',
        reviewer_name=reviewer_name,
    ))
    record.updated_at = datetime.now().isoformat()
    return record


def reject_review(record: ReviewRecord, reviewer_name: str, comment: str) -> ReviewRecord:
    """Reject the review with comments."""
    record.status = 'rejected'
    record.history.append(ReviewAction(
        action='reject',
        comment=comment,
        reviewer_name=reviewer_name,
    ))
    record.updated_at = datetime.now().isoformat()
    return record


def request_revision(record: ReviewRecord, reviewer_name: str, comment: str) -> ReviewRecord:
    """Request revision of the report."""
    record.status = 'revision_needed'
    record.history.append(ReviewAction(
        action='revise',
        comment=comment,
        reviewer_name=reviewer_name,
    ))
    record.updated_at = datetime.now().isoformat()
    return record


def save_review(record: ReviewRecord) -> str:
    """Save review record to disk."""
    _ensure_reviews_dir()
    path = os.path.join(REVIEWS_DIR, f"{record.review_id}.json")
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(record.to_dict(), f, ensure_ascii=False, indent=2)
    return path


def load_review(review_id: str) -> ReviewRecord | None:
    """Load review record from disk."""
    path = os.path.join(REVIEWS_DIR, f"{review_id}.json")
    if not os.path.exists(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return ReviewRecord.from_dict(data)


def list_reviews(status: str = '') -> list[dict]:
    """List all review records, optionally filtered by status."""
    _ensure_reviews_dir()
    reviews = []
    for fname in os.listdir(REVIEWS_DIR):
        if fname.endswith('.json'):
            try:
                with open(os.path.join(REVIEWS_DIR, fname), 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if status and data.get('status') != status:
                    continue
                reviews.append(data)
            except (json.JSONDecodeError, IOError):
                continue
    reviews.sort(key=lambda r: r.get('updated_at', ''), reverse=True)
    return reviews
