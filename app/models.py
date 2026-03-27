from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Text, JSON
from sqlalchemy.orm import relationship
from datetime import datetime, timezone

from app.database import Base


class Deal(Base):
    __tablename__ = "deals"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(500), nullable=False)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    comparisons = relationship("Comparison", back_populates="deal", cascade="all, delete-orphan")


class Comparison(Base):
    __tablename__ = "comparisons"

    id = Column(Integer, primary_key=True, autoincrement=True)
    deal_id = Column(Integer, ForeignKey("deals.id"), nullable=False)
    word_filename = Column(String(500), nullable=False)
    excel_filename = Column(String(500), nullable=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    status = Column(String(20), nullable=True)  # "pass", "fail", or None (pending)
    output_filename = Column(String(500), nullable=True)

    # Store parsed tables as JSON
    parsed_tables = Column(JSON, nullable=True)  # list of 2D arrays from Word
    selected_table_indices = Column(JSON, nullable=True)  # list of selected table indices
    detected_ranges = Column(JSON, nullable=True)  # dict: tab_index -> range string
    user_range_overrides = Column(JSON, nullable=True)  # dict: tab_index -> range string

    deal = relationship("Deal", back_populates="comparisons")
    comparison_tables = relationship("ComparisonTable", back_populates="comparison", cascade="all, delete-orphan")


class ComparisonTable(Base):
    __tablename__ = "comparison_tables"

    id = Column(Integer, primary_key=True, autoincrement=True)
    comparison_id = Column(Integer, ForeignKey("comparisons.id"), nullable=False)
    table_index = Column(Integer, nullable=False)
    table_label = Column(String(200), nullable=True)
    excel_tab_name = Column(String(200), nullable=True)
    detected_range = Column(String(50), nullable=True)
    user_range_override = Column(String(50), nullable=True)
    precision_overrides = Column(JSON, nullable=True)  # dict: row_index -> precision int
    match_count = Column(Integer, nullable=True)
    mismatch_count = Column(Integer, nullable=True)
    total_cells = Column(Integer, nullable=True)
    comparison_data = Column(JSON, nullable=True)  # store the X-Y grid for display

    comparison = relationship("Comparison", back_populates="comparison_tables")
