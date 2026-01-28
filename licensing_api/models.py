from __future__ import annotations

from datetime import datetime, timedelta
from sqlalchemy import (
    Column,
    Integer,
    String,
    Boolean,
    DateTime,
    ForeignKey,
    UniqueConstraint,
)
from sqlalchemy.orm import relationship

from licensing_api.database import Base


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), unique=True, index=True, nullable=False)
    password_hash = Column(String(255), nullable=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(
        DateTime, default=datetime.utcnow, onupdate=datetime.utcnow
    )

    licenses = relationship("License", back_populates="user", cascade="all,delete")


class License(Base):
    __tablename__ = "licenses"
    __table_args__ = (
        UniqueConstraint("code", name="uq_license_code"),
    )

    id = Column(Integer, primary_key=True, index=True)
    code = Column(String(64), nullable=False)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    device_fingerprint = Column(String(255), nullable=True)
    is_active = Column(Boolean, default=True)
    activated_at = Column(DateTime, nullable=True)
    expires_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    user = relationship("User", back_populates="licenses")
    devices = relationship(
        "LicenseDevice",
        back_populates="license",
        cascade="all,delete-orphan",
    )

    def mark_activated(self, fingerprint: str, validity_days: int | None = None):
        self.device_fingerprint = fingerprint
        self.activated_at = datetime.utcnow()
        if validity_days:
            self.expires_at = self.activated_at + timedelta(days=validity_days)


class LicenseDevice(Base):
    __tablename__ = "license_devices"
    __table_args__ = (
        UniqueConstraint("license_id", "fingerprint", name="uq_license_device"),
    )

    id = Column(Integer, primary_key=True, index=True)
    license_id = Column(Integer, ForeignKey("licenses.id"), nullable=False)
    fingerprint = Column(String(255), nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    license = relationship("License", back_populates="devices")
