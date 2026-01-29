from __future__ import annotations

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, EmailStr, constr


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"


class LoginRequest(BaseModel):
    email: EmailStr
    password: constr(min_length=6)


class LicenseActivationRequest(BaseModel):
    code: str
    fingerprint: constr(min_length=6)


class LicenseValidationRequest(BaseModel):
    fingerprint: constr(min_length=6)


class LicenseDeviceInfo(BaseModel):
    fingerprint: str
    created_at: Optional[datetime]

    class Config:
        orm_mode = True


class LicenseInfo(BaseModel):
    code: str
    is_active: bool
    activated_at: Optional[datetime]
    expires_at: Optional[datetime]
    device_fingerprint: Optional[str]
    devices: list[LicenseDeviceInfo] = []

    class Config:
        orm_mode = True


class UserInfo(BaseModel):
    email: EmailStr
    is_active: bool
    licenses: list[LicenseInfo] = []

    class Config:
        orm_mode = True


class AdminCreateUserLicenseRequest(BaseModel):
    email: EmailStr
    password: constr(min_length=6)
    code: str
    expires_at: Optional[datetime] = None
