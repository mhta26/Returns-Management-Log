from __future__ import annotations
import re
from datetime import date, datetime
from zoneinfo import ZoneInfo
from pydantic import BaseModel, Field, ConfigDict
from typing import Annotated


def now():
    return datetime.now(ZoneInfo('Asia/Riyadh'))


def normalized(text):
    text = str(text).strip().lower().translate(str.maketrans('أإآىةـ', 'ااايه '))
    return re.sub(r'[\u064b-\u065f\u0670]', '', text)


def status_for(expiry: date, receipt_date: date, settings: dict):
    remaining = (expiry - receipt_date).days
    thresholds = settings['thresholds']
    index = next((i for i, limit in enumerate(thresholds) if remaining <= limit), 3)
    return settings['labels'][index], index, remaining


class Line(BaseModel):
    model_config = ConfigDict(extra='forbid')
    product_code: str = Field(min_length=1, max_length=100)
    cartons: Annotated[int, Field(strict=True, ge=0, le=1000000000)] = 0
    boxes: Annotated[int, Field(strict=True, ge=0, le=1000000000)] = 0
    pieces: Annotated[int, Field(strict=True, ge=0, le=1000000000)] = 0
    expiry: date
    reason: str = Field(default='', max_length=500)
    notes: str = Field(default='', max_length=2000)


class ReceiptInput(BaseModel):
    model_config = ConfigDict(extra='forbid')
    request_id: str = Field(min_length=20, max_length=100)
    operator: str = Field(min_length=1, max_length=200)
    customer_code: str = Field(min_length=1, max_length=100)
    driver_code: str = Field(default='', max_length=100)
    return_number: str = Field(default='', max_length=200)
    invoice_date: date | None = None
    notes: str = Field(default='', max_length=2000)
    items: list[Line] = Field(min_length=1, max_length=220)


class SettingsInput(BaseModel):
    thresholds: list[int] = Field(min_length=3, max_length=3)
    labels: list[str] = Field(min_length=4, max_length=4)
    reasons: list[str] = Field(min_length=1, max_length=50)
    next_serial: int = Field(ge=1, le=2147483646)
    company: str = Field(min_length=1, max_length=300)
    membership: str = Field(default='', max_length=100)
    keeper: str = Field(default='', max_length=200)
    accountant: str = Field(default='', max_length=200)
    document_title: str = Field(min_length=1, max_length=200)
