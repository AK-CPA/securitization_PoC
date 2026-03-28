from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, JSON, Float
from sqlalchemy.orm import relationship
from datetime import datetime, timezone

from app.database import Base


class Deal(Base):
    __tablename__ = "deals"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(500), nullable=False)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    comparisons = relationship("Comparison", back_populates="deal", cascade="all, delete-orphan")
    loose_comparisons = relationship("LooseComparison", back_populates="deal", cascade="all, delete-orphan")


class Comparison(Base):
    __tablename__ = "comparisons"

    id = Column(Integer, primary_key=True, autoincrement=True)
    deal_id = Column(Integer, ForeignKey("deals.id"), nullable=False)
    word_filename = Column(String(500), nullable=False)
    excel_filename = Column(String(500), nullable=True)  # kept for history display (legacy)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    status = Column(String(20), nullable=True)  # "pass", "fail", or None (pending)
    output_filename = Column(String(500), nullable=True)

    # Store parsed tables as JSON
    parsed_tables = Column(JSON, nullable=True)  # list of 2D arrays from Word
    selected_table_indices = Column(JSON, nullable=True)  # list of selected table indices
    detected_ranges = Column(JSON, nullable=True)  # legacy — kept for backward compat
    user_range_overrides = Column(JSON, nullable=True)  # legacy

    deal = relationship("Deal", back_populates="comparisons")
    comparison_tables = relationship("ComparisonTable", back_populates="comparison", cascade="all, delete-orphan")
    uploaded_files = relationship("UploadedFile", back_populates="comparison", cascade="all, delete-orphan")


class UploadedFile(Base):
    """A source file (Excel or XML) uploaded for a comparison."""
    __tablename__ = "uploaded_files"

    id = Column(Integer, primary_key=True, autoincrement=True)
    comparison_id = Column(Integer, ForeignKey("comparisons.id"), nullable=False)
    filename = Column(String(500), nullable=False)
    file_type = Column(String(10), nullable=False)  # "xlsx" or "xml"
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    comparison = relationship("Comparison", back_populates="uploaded_files")
    comparison_tables = relationship("ComparisonTable", back_populates="uploaded_file")


class ComparisonTable(Base):
    __tablename__ = "comparison_tables"

    id = Column(Integer, primary_key=True, autoincrement=True)
    comparison_id = Column(Integer, ForeignKey("comparisons.id"), nullable=False)
    table_index = Column(Integer, nullable=False)
    table_label = Column(String(200), nullable=True)

    # Per-table file mapping (new: each table points to a specific uploaded file + sheet + range)
    uploaded_file_id = Column(Integer, ForeignKey("uploaded_files.id"), nullable=True)
    excel_tab_name = Column(String(200), nullable=True)
    detected_ranges = Column(JSON, nullable=True)  # list of range strings detected on the sheet
    selected_range = Column(String(50), nullable=True)  # the range the user chose (from detected or manual)
    user_range_override = Column(String(50), nullable=True)  # manual override

    # Legacy single-range fields (kept for backward compat)
    detected_range = Column(String(50), nullable=True)

    precision_overrides = Column(JSON, nullable=True)  # dict: row_index -> precision int
    match_count = Column(Integer, nullable=True)
    mismatch_count = Column(Integer, nullable=True)
    total_cells = Column(Integer, nullable=True)
    comparison_data = Column(JSON, nullable=True)  # store the X-Y grid for display

    comparison = relationship("Comparison", back_populates="comparison_tables")
    uploaded_file = relationship("UploadedFile", back_populates="comparison_tables")


# ── Loose Language Tie-Out models ─────────────────────────────────────────────

class SentenceTemplate(Base):
    """Reusable set of candidate sentences for loose language matching."""
    __tablename__ = "sentence_templates"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(500), nullable=False)
    sentences = Column(JSON, nullable=False)  # list of sentence strings
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))


class LooseComparison(Base):
    """A loose-language tie-out comparison for a deal."""
    __tablename__ = "loose_comparisons"

    id = Column(Integer, primary_key=True, autoincrement=True)
    deal_id = Column(Integer, ForeignKey("deals.id"), nullable=False)
    word_filename = Column(String(500), nullable=False)
    excel_filename = Column(String(500), nullable=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    status = Column(String(20), nullable=True)  # "pass", "fail", or None
    output_filename = Column(String(500), nullable=True)

    # Candidate sentences provided by user
    candidate_sentences = Column(JSON, nullable=True)  # list of strings

    # Full text extracted from Word document (for searching)
    document_text = Column(Text, nullable=True)

    deal = relationship("Deal", back_populates="loose_comparisons")
    loose_items = relationship("LooseComparisonItem", back_populates="loose_comparison", cascade="all, delete-orphan")


class LooseComparisonItem(Base):
    """One matched sentence and its comparison result."""
    __tablename__ = "loose_comparison_items"

    id = Column(Integer, primary_key=True, autoincrement=True)
    loose_comparison_id = Column(Integer, ForeignKey("loose_comparisons.id"), nullable=False)

    # The candidate sentence provided by user
    candidate_sentence = Column(Text, nullable=False)
    # The matched sentence found in the document
    matched_sentence = Column(Text, nullable=True)
    similarity_score = Column(Float, nullable=True)

    # Claude API extraction results
    extraction_prompt = Column(Text, nullable=True)  # what we asked Claude
    extracted_values = Column(JSON, nullable=True)  # values Claude found in Excel
    document_values = Column(JSON, nullable=True)  # values found in the sentence
    comparison_result = Column(JSON, nullable=True)  # {matches: [...], mismatches: [...]}

    status = Column(String(20), nullable=True)  # "pass", "fail", "no_match", "error"
    error_message = Column(Text, nullable=True)

    loose_comparison = relationship("LooseComparison", back_populates="loose_items")
