import json
import hashlib
import html
from io import BytesIO
import os
import secrets
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import wraps
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import urljoin, urlsplit
from zoneinfo import ZoneInfo
from zipfile import BadZipFile, ZipFile

import click
from flask import (
    Flask,
    abort,
    current_app,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from flask_sqlalchemy import SQLAlchemy
from flask_wtf import FlaskForm
from flask_wtf.csrf import CSRFProtect
import pycountry
from airport_lookup import AirportLookupError, airport_suggestions
from postal_code_lookup import PostalCodeLookupError, postal_code_suggestions
from sqlalchemy import UniqueConstraint, inspect, or_, select, text, update
from werkzeug.security import check_password_hash, generate_password_hash
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.platypus import Image as PDFImage
from wtforms import (
    BooleanField,
    DateField,
    DecimalField,
    FieldList,
    Form,
    FormField,
    PasswordField,
    SelectField,
    StringField,
    SubmitField,
    TextAreaField,
    TimeField,
)
from wtforms.fields import SelectChoice
from wtforms.validators import DataRequired, EqualTo, Length, NumberRange, Optional
from werkzeug.utils import secure_filename
from postal_code_lookup import PostalCodeLookupError, postal_code_suggestions


BASE_DIR = Path(__file__).resolve().parent
db = SQLAlchemy()
login_manager = LoginManager()
csrf = CSRFProtect()
login_manager.login_view = "login"
login_manager.login_message = "Please sign in to continue."
login_manager.login_message_category = "warning"

SHIPMENT_STAGES = [
    "Received",
    "Document Check",
    "Customs Filed",
    "Customs Cleared",
    "Duty Paid",
    "Warehouse",
    "Release to Client",
    "Re-export",
    "Closed",
]
CUSTOMS_STATUSES = ["Not Filed", "Filed", "Cleared", "Hold", "Rejected"]
AIR_CUSTOMS_STATUSES = ["Not Filed", "Filed", "Cleared", "Hold"]
PAYMENT_STATUSES = ["Pending", "Partially Paid", "Paid", "Overdue"]
RE_EXPORT_STATUSES = ["Not Planned", "Planned", "In Progress", "Completed"]
TRACKING_STATUSES = ["In Transit", "Out for Delivery", "Delivered", "Delayed"]
TRANSPORT_TYPES = ["Air", "Road", "Sea", "No transport", "Other"]
TRANSPORT_REFERENCE_CODES = {
    "Air": "A",
    "Road": "R",
    "Sea": "S",
    "No transport": "N",
    "Other": "O",
}
GOODS_DIRECTIONS = ("export", "import")
GOODS_FORM_FIELDS = (
    "marking",
    "pieces",
    "length_cm",
    "width_cm",
    "height_cm",
    "pack_type",
    "pack_other",
    "weight_kg",
    "description",
    "clearance_method",
)
PACKING_TYPES = [
    "Wooden Packing",
    "Carton",
    "Crate",
    "Pallet",
    "Bag",
    "Drum",
    "Bundle",
    "Coil",
    "Packet",
    "Machine",
    "Other",
]
CLEARANCE_METHODS = [
    "Temporary Import",
    "Temporary Export",
    "Definitive Export",
    "ATA",
    "Definitive Import",
    "FTWZ Sales",
    "Ex-Bond",
    "Re-Import",
    "Re-export",
]
FILE_REFERENCE_CODES = ["A", "R", "S", "N", "O", "C"]
ON_SITE_HANDLING_CHOICES = ["Yes", "No", "Local"]
CONTACT_KINDS = ["exhibitor", "agent", "customer", "payment", "pickup"]
EVENT_FILE_CONTACT_KINDS = ["exhibitor", "agent", "customer"]
AIR_SHIPMENT_FORM_FIELDS = (
    "mawb_no",
    "hawb_no",
    "hawb_departure",
    "hawb_arrival",
    "hawb_departure_date",
    "hawb_arrival_date",
    "tracking_status",
    "customs_status",
    "payment_status",
    "stage",
)
DEPARTMENTS = [
    {"slug": "nom", "name": "NOM"},
    {"slug": "general", "name": "GENERAL"},
    {"slug": "defense", "name": "DEFENSE"},
    {"slug": "import-export", "name": "IMPORT EXPORT", "short_name": "IIM"},
    {"slug": "ftwz", "name": "FTWZ"},
    {"slug": "local", "name": "LOCAL"},
]
IDLE_TIMEOUT = timedelta(minutes=15)
ACTIVE_WINDOW = timedelta(minutes=2)
HEARTBEAT_INTERVAL = timedelta(seconds=30)
DISPLAY_TIMEZONE = ZoneInfo("Asia/Kolkata")
EXPORT_FIELDS = [
    ("reference", "Shipment / AWB / BL"),
    ("client", "Client / company"),
    ("agent_name", "Agent name"),
    ("exhibitor_name", "Exhibitor name"),
    ("exhibition_name", "Exhibition name"),
    ("supplier", "Supplier / exporter"),
    ("origin", "Origin"),
    ("destination", "Destination"),
    ("tracking_status", "Tracking status"),
    ("arrival_date", "Arrival date"),
    ("arrival_time", "Arrival time"),
    ("expected_reexport_date", "Expected re-export date"),
    ("shipment_value", "Shipment value"),
    ("stage", "Operational stage"),
    ("customs_status", "Customs status"),
    ("customs_reference", "Customs reference"),
    ("payment_status", "Payment status"),
    ("pending_items", "Pending items"),
    ("re_export_status", "Re-export status"),
    ("re_export_reference", "Re-export reference"),
    ("container_no", "Container / seal"),
    ("warehouse_location", "Warehouse location"),
    ("notes", "Notes"),
    ("updated_at", "Last updated (IST)"),
]
EVENT_EXPORT_FIELDS = [
    ("master_file_ref", "Master file reference"),
    ("exhibition_name", "Exhibition name"),
    ("start_date", "Exhibition start date"),
    ("end_date", "Exhibition end date"),
    ("venue", "Venue"),
    ("organiser_name", "Organiser name"),
    ("move_in_start", "Move-in start date"),
    ("move_in_end", "Move-in end date"),
    ("move_out_start", "Move-out start date"),
    ("move_out_end", "Move-out end date"),
    ("updated_at", "Last updated (IST)"),
]
EVENT_FILE_EXPORT_FIELDS = [
    ("id", "File number"),
    ("file_reference", "File reference"),
    ("file_date", "Date"),
    ("event_name", "Event name"),
    ("transport_type", "Type of transport"),
    ("custom_transport", "Other transport details"),
    ("on_site_handling", "On-site handling"),
    ("hall_number", "Hall number"),
    ("booth_number", "Booth number"),
    ("exhibitor_name", "Exhibitor / company"),
    ("exhibitor_contact_person", "Exhibitor contact person"),
    ("exhibitor_email", "Exhibitor email"),
    ("exhibitor_phone", "Exhibitor phone"),
    ("agent_name", "Agent / company"),
    ("agent_contact_person", "Agent contact person"),
    ("agent_email", "Agent email"),
    ("agent_phone", "Agent phone"),
    ("customer_name", "Customer / company"),
    ("customer_contact_person", "Customer contact person"),
    ("customer_email", "Customer email"),
    ("customer_phone", "Customer phone"),
    ("payment_name", "Payment company"),
    ("payment_contact_person", "Payment contact person"),
    ("payment_email", "Payment email"),
    ("payment_phone", "Payment phone"),
    ("pickup_name", "Pickup company"),
    ("pickup_contact_person", "Pickup contact person"),
    ("pickup_email", "Pickup email"),
    ("pickup_phone", "Pickup phone"),
    ("tax_formalities", "Tax formalities"),
    ("insurance", "Insurance"),
    ("insurance_value", "Insurance value (INR)"),
    ("updated_at", "Last updated (IST)"),
]
AGENT_EXPORT_FIELDS = [
    ("company_name", "Company name"),
    ("country", "Country"),
    ("address", "Address"),
    ("city", "City"),
    ("postal_code", "ZIP / Postal code"),
    ("company_phone", "Company phone"),
    ("company_email", "Company email"),
    ("vat_number", "VAT number"),
    ("contact_people", "Contact persons"),
    ("updated_at", "Last updated (IST)"),
]
COUNTRY_OPTIONS = sorted(
    [
        {"name": country.name, "code": country.alpha_2}
        for country in pycountry.countries
    ],
    key=lambda country: country["name"].casefold(),
)
CUSTOMER_EXPORT_FIELDS = [
    ("company_name", "Customer / company"),
    ("contact_person_name", "Contact person"),
    ("address", "Address"),
    ("email", "Email"),
    ("phone", "Phone"),
    ("updated_at", "Last updated (IST)"),
]


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    is_admin = db.Column(db.Boolean, nullable=False, default=False)
    can_view = db.Column(db.Boolean, nullable=False, default=True, server_default=db.true())
    can_add = db.Column(db.Boolean, nullable=False, default=False, server_default=db.false())
    can_edit = db.Column(db.Boolean, nullable=False, default=False, server_default=db.false())
    can_delete = db.Column(db.Boolean, nullable=False, default=False, server_default=db.false())

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def has_permission(self, permission):
        return self.is_admin or bool(getattr(self, f"can_{permission}", False))


class UserSession(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    token_hash = db.Column(db.String(64), unique=True, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    username = db.Column(db.String(80), nullable=False)
    started_at = db.Column(db.DateTime(timezone=True), nullable=False)
    last_seen_at = db.Column(db.DateTime(timezone=True), nullable=False)
    ended_at = db.Column(db.DateTime(timezone=True), nullable=True)
    end_reason = db.Column(db.String(40), nullable=False, default="")


class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    actor_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True, index=True)
    actor_username = db.Column(db.String(80), nullable=False)
    action = db.Column(db.String(40), nullable=False)
    entity_type = db.Column(db.String(40), nullable=False)
    entity_id = db.Column(db.Integer, nullable=True)
    entity_label = db.Column(db.String(200), nullable=False, default="")
    details_json = db.Column(db.Text, nullable=False, default="{}")
    created_at = db.Column(db.DateTime(timezone=True), nullable=False)

class Shipment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    reference = db.Column(db.String(120), nullable=False)
    reference_key = db.Column(db.String(120), unique=True, nullable=False)
    agent_name = db.Column(db.String(200), nullable=False, default="")
    client = db.Column(db.String(200), nullable=False)
    exhibitor_name = db.Column(db.String(200), nullable=False, default="")
    exhibition_name = db.Column(db.String(200), nullable=False, default="")
    supplier = db.Column(db.String(200), nullable=False, default="")
    origin = db.Column(db.String(200), nullable=False, default="")
    destination = db.Column(db.String(200), nullable=False, default="")
    tracking_status = db.Column(
        db.String(40),
        nullable=False,
        default="In Transit",
        server_default="In Transit",
    )
    arrival_date = db.Column(db.Date, nullable=False)
    arrival_time = db.Column(db.Time, nullable=True)
    expected_reexport_date = db.Column(db.Date, nullable=True)
    shipment_value = db.Column(db.String(100), nullable=False, default="")
    stage = db.Column(db.String(100), nullable=False, default="Received")
    customs_status = db.Column(db.String(100), nullable=False, default="Not Filed")
    customs_reference = db.Column(db.String(200), nullable=False, default="")
    payment_status = db.Column(db.String(100), nullable=False, default="Pending")
    pending_items = db.Column(db.Text, nullable=False, default="")
    re_export_status = db.Column(db.String(100), nullable=False, default="Not Planned")
    re_export_reference = db.Column(db.String(200), nullable=False, default="")
    container_no = db.Column(db.String(200), nullable=False, default="")
    warehouse_location = db.Column(db.String(200), nullable=False, default="")
    notes = db.Column(db.Text, nullable=False, default="")
    updated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=True,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )


class Event(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    department = db.Column(db.String(40), nullable=False, index=True)
    exhibition_name = db.Column(db.String(200), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    venue = db.Column(db.String(200), nullable=False)
    organiser_name = db.Column(db.String(200), nullable=False)
    move_in_start = db.Column(db.Date, nullable=True)
    move_in_end = db.Column(db.Date, nullable=True)
    move_out_start = db.Column(db.Date, nullable=True)
    move_out_end = db.Column(db.Date, nullable=True)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )
    reference_record = db.relationship(
        "EventMasterReference",
        back_populates="event",
        uselist=False,
        cascade="all, delete-orphan",
    )
    file_sequences = db.relationship(
        "EventFileSequence",
        back_populates="event",
        cascade="all, delete-orphan",
    )

    @property
    def master_file_ref(self):
        return self.reference_record.master_ref if self.reference_record else ""

    @master_file_ref.setter
    def master_file_ref(self, value):
        master_ref = (value or "").strip()
        if self.reference_record is None:
            self.reference_record = EventMasterReference(
                master_ref=master_ref,
                master_ref_key=master_ref.casefold(),
            )
        else:
            self.reference_record.master_ref = master_ref
            self.reference_record.master_ref_key = master_ref.casefold()


class EventMasterReference(db.Model):
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), primary_key=True)
    master_ref = db.Column(db.String(50), nullable=False)
    master_ref_key = db.Column(db.String(50), nullable=False, unique=True)
    event = db.relationship("Event", back_populates="reference_record")


class EventFileSequence(db.Model):
    __table_args__ = (
        UniqueConstraint("event_id", "transport_code", name="uq_event_file_sequence"),
    )

    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False, index=True)
    transport_code = db.Column(db.String(1), nullable=False)
    next_sequence = db.Column(db.Integer, nullable=False, default=1)
    event = db.relationship("Event", back_populates="file_sequences")


class Contact(db.Model):
    __table_args__ = (
        UniqueConstraint("kind", "name_key", name="uq_contact_kind_name"),
    )

    id = db.Column(db.Integer, primary_key=True)
    kind = db.Column(db.String(20), nullable=False, index=True)
    name_key = db.Column(db.String(200), nullable=False)
    company_name = db.Column(db.String(200), nullable=False)
    contact_person_name = db.Column(db.String(200), nullable=False, default="")
    address = db.Column(db.Text, nullable=False, default="")
    email = db.Column(db.String(254), nullable=False, default="")
    phone = db.Column(db.String(50), nullable=False, default="")
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )


class AgentProfile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    contact_id = db.Column(db.Integer, db.ForeignKey("contact.id"), nullable=False, unique=True)
    country = db.Column(db.String(100), nullable=False, default="")
    address = db.Column(db.Text, nullable=False, default="")
    city = db.Column(db.String(150), nullable=False, default="")
    postal_code = db.Column(db.String(30), nullable=False, default="")
    company_phone = db.Column(db.String(50), nullable=False, default="")
    company_email = db.Column(db.String(254), nullable=False, default="")
    vat_number = db.Column(db.String(100), nullable=False, default="")
    contact = db.relationship("Contact", backref=db.backref("agent_profile", uselist=False))
    contact_people = db.relationship(
        "AgentContactPerson",
        back_populates="profile",
        cascade="all, delete-orphan",
        order_by="AgentContactPerson.id",
    )


class AgentContactPerson(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    profile_id = db.Column(db.Integer, db.ForeignKey("agent_profile.id"), nullable=False, index=True)
    contact_name = db.Column(db.String(200), nullable=False)
    phone = db.Column(db.String(50), nullable=False)
    profile = db.relationship("AgentProfile", back_populates="contact_people")


class EventShipmentFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False, index=True)
    transport_type = db.Column(db.String(30), nullable=False)
    custom_transport = db.Column(db.String(100), nullable=False, default="")
    on_site_handling = db.Column(db.String(20), nullable=False)
    exhibitor_contact_id = db.Column(db.Integer, db.ForeignKey("contact.id"), nullable=True)
    agent_contact_id = db.Column(db.Integer, db.ForeignKey("contact.id"), nullable=True)
    customer_contact_id = db.Column(db.Integer, db.ForeignKey("contact.id"), nullable=True)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )
    event = db.relationship("Event")
    exhibitor_contact = db.relationship("Contact", foreign_keys=[exhibitor_contact_id])
    agent_contact = db.relationship("Contact", foreign_keys=[agent_contact_id])
    customer_contact = db.relationship("Contact", foreign_keys=[customer_contact_id])
    details = db.relationship(
        "EventShipmentFileDetails",
        back_populates="file",
        uselist=False,
        cascade="all, delete-orphan",
    )
    reference_record = db.relationship(
        "EventShipmentFileReference",
        back_populates="file",
        uselist=False,
        cascade="all, delete-orphan",
    )
    goods_items = db.relationship(
        "EventShipmentGood",
        back_populates="file",
        cascade="all, delete-orphan",
        order_by="EventShipmentGood.direction, EventShipmentGood.sort_order, EventShipmentGood.id",
    )
    proposals = db.relationship(
        "EventFileProposal",
        back_populates="file",
        cascade="all, delete-orphan",
        order_by="EventFileProposal.id.desc()",
    )
    delivery_note = db.relationship(
        "EventFileDeliveryNote",
        back_populates="file",
        uselist=False,
        cascade="all, delete-orphan",
    )
    delivery_note_history = db.relationship(
        "EventFileDeliveryNoteHistory",
        back_populates="file",
        cascade="all, delete-orphan",
        order_by="EventFileDeliveryNoteHistory.generated_at.desc(), EventFileDeliveryNoteHistory.id.desc()",
    )
    air_details = db.relationship(
        "EventShipmentFileAirDetails",
        back_populates="file",
        uselist=False,
        cascade="all, delete-orphan",
    )

    @property
    def file_date(self):
        if self.details:
            return self.details.file_date
        return self.created_at.date() if self.created_at else None

    @property
    def file_reference(self):
        return self.reference_record.reference if self.reference_record else ""

    @property
    def event_name(self):
        return self.event.exhibition_name

    @property
    def hall_number(self):
        return self.details.hall_number if self.details else ""

    @property
    def booth_number(self):
        return self.details.booth_number if self.details else ""

    @property
    def tax_formalities(self):
        return self.details.tax_formalities if self.details else ""

    @property
    def insurance(self):
        return self.details.insurance if self.details else ""

    @property
    def insurance_value(self):
        return self.details.insurance_value if self.details else Decimal("0.00")

    @property
    def payment_name(self):
        return self.details.payment_contact.company_name if self.details and self.details.payment_contact else ""

    @property
    def payment_contact_person(self):
        return self.details.payment_contact.contact_person_name if self.details and self.details.payment_contact else ""

    @property
    def payment_email(self):
        return self.details.payment_contact.email if self.details and self.details.payment_contact else ""

    @property
    def payment_phone(self):
        return self.details.payment_contact.phone if self.details and self.details.payment_contact else ""

    @property
    def pickup_name(self):
        return self.details.pickup_contact.company_name if self.details and self.details.pickup_contact else ""

    @property
    def pickup_contact_person(self):
        return self.details.pickup_contact.contact_person_name if self.details and self.details.pickup_contact else ""

    @property
    def pickup_email(self):
        return self.details.pickup_contact.email if self.details and self.details.pickup_contact else ""

    @property
    def pickup_phone(self):
        return self.details.pickup_contact.phone if self.details and self.details.pickup_contact else ""

    @property
    def exhibitor_name(self):
        return self.exhibitor_contact.company_name

    @property
    def exhibitor_contact_person(self):
        return self.exhibitor_contact.contact_person_name

    @property
    def exhibitor_email(self):
        return self.exhibitor_contact.email

    @property
    def exhibitor_phone(self):
        return self.exhibitor_contact.phone

    @property
    def agent_name(self):
        return self.agent_contact.company_name

    @property
    def agent_contact_person(self):
        return self.agent_contact.contact_person_name

    @property
    def agent_email(self):
        return self.agent_contact.email

    @property
    def agent_phone(self):
        return self.agent_contact.phone

    @property
    def customer_name(self):
        return self.customer_contact.company_name

    @property
    def customer_contact_person(self):
        return self.customer_contact.contact_person_name

    @property
    def customer_email(self):
        return self.customer_contact.email

    @property
    def customer_phone(self):
        return self.customer_contact.phone


class EventShipmentFileDetails(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey("event_shipment_file.id"), nullable=False, unique=True)
    file_date = db.Column(db.Date, nullable=False, default=date.today)
    payment_contact_id = db.Column(db.Integer, db.ForeignKey("contact.id"), nullable=True)
    pickup_contact_id = db.Column(db.Integer, db.ForeignKey("contact.id"), nullable=True)
    hall_number = db.Column(db.String(100), nullable=False, default="")
    booth_number = db.Column(db.String(100), nullable=False, default="")
    tax_formalities = db.Column(db.String(20), nullable=False, default="")
    insurance = db.Column(db.String(20), nullable=False, default="")
    insurance_value = db.Column(db.Numeric(12, 2), nullable=False, default=Decimal("0.00"))
    file = db.relationship("EventShipmentFile", back_populates="details")
    payment_contact = db.relationship("Contact", foreign_keys=[payment_contact_id])
    pickup_contact = db.relationship("Contact", foreign_keys=[pickup_contact_id])


class EventShipmentFileAirDetails(db.Model):
    file_id = db.Column(db.Integer, db.ForeignKey("event_shipment_file.id"), primary_key=True)
    mawb_no = db.Column(db.String(50), nullable=False, default="")
    hawb_no = db.Column(db.String(50), nullable=False, default="")
    hawb_departure = db.Column(db.String(100), nullable=False, default="")
    hawb_arrival = db.Column(db.String(100), nullable=False, default="")
    hawb_departure_date = db.Column(db.Date, nullable=True)
    hawb_arrival_date = db.Column(db.Date, nullable=True)
    tracking_status = db.Column(db.String(40), nullable=False, default="In Transit")
    customs_status = db.Column(db.String(100), nullable=False, default="Not Filed")
    payment_status = db.Column(db.String(40), nullable=False, default="Pending")
    stage = db.Column(db.String(100), nullable=False, default="Received")
    file = db.relationship("EventShipmentFile", back_populates="air_details")


class EventShipmentGood(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey("event_shipment_file.id"), nullable=False, index=True)
    direction = db.Column(db.String(10), nullable=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    marking = db.Column(db.String(120), nullable=False)
    pieces = db.Column(db.Integer, nullable=False)
    length_cm = db.Column(db.Numeric(12, 2), nullable=False)
    width_cm = db.Column(db.Numeric(12, 2), nullable=False)
    height_cm = db.Column(db.Numeric(12, 2), nullable=False)
    pack_type = db.Column(db.String(60), nullable=False)
    pack_other = db.Column(db.String(120), nullable=False, default="")
    weight_kg = db.Column(db.Numeric(12, 2), nullable=False, default=Decimal("0.00"))
    description = db.Column(db.String(500), nullable=False, default="")
    clearance_method = db.Column(db.String(60), nullable=False)
    file = db.relationship("EventShipmentFile", back_populates="goods_items")


class EventShipmentFileReference(db.Model):
    __table_args__ = (
        UniqueConstraint(
            "event_id",
            "transport_code",
            "sequence",
            name="uq_event_transport_sequence",
        ),
    )

    file_id = db.Column(db.Integer, db.ForeignKey("event_shipment_file.id"), primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False, index=True)
    transport_code = db.Column(db.String(1), nullable=False)
    sequence = db.Column(db.Integer, nullable=False)
    reference = db.Column(db.String(120), nullable=False, unique=True)
    file = db.relationship("EventShipmentFile", back_populates="reference_record")


class EventFileProposal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(db.Integer, db.ForeignKey("event_shipment_file.id"), nullable=False, index=True)
    quote_number = db.Column(db.String(40), nullable=True, unique=True)
    recipient_type = db.Column(db.String(20), nullable=False)
    recipient_name = db.Column(db.String(200), nullable=False)
    recipient_contact = db.Column(db.String(200), nullable=False, default="")
    recipient_phone = db.Column(db.String(50), nullable=False, default="")
    recipient_email = db.Column(db.String(254), nullable=False, default="")
    recipient_address = db.Column(db.Text, nullable=False, default="")
    currency = db.Column(db.String(3), nullable=False, default="INR")
    issue_date = db.Column(db.Date, nullable=False)
    valid_until = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(30), nullable=False, default="DRAFT")
    discount_type = db.Column(db.String(10), nullable=False, default="percent")
    discount_value = db.Column(db.Numeric(16, 2), nullable=False, default=Decimal("0.00"))
    discount_percent = db.Column(db.Numeric(5, 2), nullable=False, default=Decimal("0.00"))
    tax_name = db.Column(db.String(60), nullable=False, default="")
    tax_percent = db.Column(db.Numeric(5, 2), nullable=False, default=Decimal("0.00"))
    subtotal = db.Column(db.Numeric(16, 2), nullable=False, default=Decimal("0.00"))
    discount_amount = db.Column(db.Numeric(16, 2), nullable=False, default=Decimal("0.00"))
    tax_amount = db.Column(db.Numeric(16, 2), nullable=False, default=Decimal("0.00"))
    total = db.Column(db.Numeric(16, 2), nullable=False, default=Decimal("0.00"))
    terms = db.Column(db.Text, nullable=False, default="")
    notes = db.Column(db.Text, nullable=False, default="")
    uploaded_filename = db.Column(db.String(255), nullable=False, default="")
    uploaded_mimetype = db.Column(db.String(120), nullable=False, default="")
    uploaded_data = db.Column(db.LargeBinary, nullable=True)
    bank_account_slot = db.Column(db.Integer, nullable=True)
    bank_label = db.Column(db.String(80), nullable=False, default="")
    bank_beneficiary = db.Column(db.String(200), nullable=False, default="")
    bank_name = db.Column(db.String(200), nullable=False, default="")
    bank_account_number = db.Column(db.String(100), nullable=False, default="")
    bank_account_type = db.Column(db.String(40), nullable=False, default="")
    bank_currency = db.Column(db.String(3), nullable=False, default="")
    bank_branch = db.Column(db.String(200), nullable=False, default="")
    bank_ifsc = db.Column(db.String(30), nullable=False, default="")
    bank_swift = db.Column(db.String(20), nullable=False, default="")
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=True,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )
    file = db.relationship("EventShipmentFile", back_populates="proposals")
    items = db.relationship(
        "EventFileProposalItem",
        back_populates="proposal",
        cascade="all, delete-orphan",
        order_by="EventFileProposalItem.sort_order, EventFileProposalItem.id",
    )


class EventFileDeliveryNote(db.Model):
    file_id = db.Column(
        db.Integer,
        db.ForeignKey("event_shipment_file.id"),
        primary_key=True,
    )
    delivery_date = db.Column(db.Date, nullable=False)
    condition_at_delivery = db.Column(db.Text, nullable=False, default="")
    receiver_name = db.Column(db.String(200), nullable=False, default="")
    receiver_designation = db.Column(db.String(120), nullable=False, default="")
    receiver_phone = db.Column(db.String(50), nullable=False, default="")
    receiver_email = db.Column(db.String(254), nullable=False, default="")
    supervisor_name = db.Column(db.String(200), nullable=False, default="")
    supervisor_designation = db.Column(db.String(120), nullable=False, default="")
    supervisor_phone = db.Column(db.String(50), nullable=False, default="")
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )
    file = db.relationship("EventShipmentFile", back_populates="delivery_note")


class EventFileDeliveryNoteHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    file_id = db.Column(
        db.Integer,
        db.ForeignKey("event_shipment_file.id"),
        nullable=False,
        index=True,
    )
    delivery_date = db.Column(db.Date, nullable=False)
    condition_at_delivery = db.Column(db.Text, nullable=False, default="")
    receiver_name = db.Column(db.String(200), nullable=False, default="")
    receiver_designation = db.Column(db.String(120), nullable=False, default="")
    receiver_phone = db.Column(db.String(50), nullable=False, default="")
    receiver_email = db.Column(db.String(254), nullable=False, default="")
    supervisor_name = db.Column(db.String(200), nullable=False, default="")
    supervisor_designation = db.Column(db.String(120), nullable=False, default="")
    supervisor_phone = db.Column(db.String(50), nullable=False, default="")
    created_by_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    generated_at = db.Column(
        db.DateTime(timezone=True),
        nullable=False,
        default=lambda: datetime.now(timezone.utc),
    )
    file = db.relationship("EventShipmentFile", back_populates="delivery_note_history")


class EventFileProposalItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    proposal_id = db.Column(
        db.Integer,
        db.ForeignKey("event_file_proposal.id"),
        nullable=False,
        index=True,
    )
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    description = db.Column(db.String(250), nullable=False)
    quantity = db.Column(db.Numeric(12, 3), nullable=False)
    unit = db.Column(db.String(40), nullable=False, default="")
    unit_price = db.Column(db.Numeric(16, 4), nullable=False)
    line_total = db.Column(db.Numeric(16, 2), nullable=False)
    proposal = db.relationship("EventFileProposal", back_populates="items")


class ProposalBankAccount(db.Model):
    slot = db.Column(db.Integer, primary_key=True)
    label = db.Column(db.String(80), nullable=False, default="")
    beneficiary = db.Column(db.String(200), nullable=False, default="")
    bank_name = db.Column(db.String(200), nullable=False, default="")
    account_number = db.Column(db.String(100), nullable=False, default="")
    account_type = db.Column(db.String(40), nullable=False, default="")
    currency = db.Column(db.String(3), nullable=False, default="")
    branch = db.Column(db.String(200), nullable=False, default="")
    ifsc = db.Column(db.String(30), nullable=False, default="")
    swift = db.Column(db.String(20), nullable=False, default="")

    @property
    def is_configured(self):
        return bool(self.beneficiary.strip() and self.bank_name.strip() and self.account_number.strip())


class AppSetting(db.Model):
    key = db.Column(db.String(100), primary_key=True)
    value = db.Column(db.String(255), nullable=False)


class LoginForm(FlaskForm):
    username = StringField("Username", validators=[DataRequired(), Length(max=80)])
    password = PasswordField("Password", validators=[DataRequired()])
    submit = SubmitField("Login")


class ShipmentForm(FlaskForm):
    reference = StringField("Shipment / AWB / BL number", validators=[DataRequired(), Length(max=120)])
    client = StringField("Client / company", validators=[DataRequired(), Length(max=200)])
    agent_name = StringField("Agent name", validators=[Optional(), Length(max=200)])
    exhibitor_name = StringField("Exhibitor name", validators=[Optional(), Length(max=200)])
    exhibition_name = StringField("Exhibition name", validators=[Optional(), Length(max=200)])
    supplier = StringField("Supplier / exporter", validators=[Optional(), Length(max=200)])
    origin = StringField("From (origin city / port)", validators=[Optional(), Length(max=200)])
    destination = StringField("To (destination city / port)", validators=[Optional(), Length(max=200)])
    arrival_date = DateField("Arrival date", format="%Y-%m-%d", validators=[DataRequired()])
    arrival_time = TimeField("Arrival time (optional)", format="%H:%M", validators=[Optional()])
    expected_reexport_date = DateField(
        "Expected re-export date",
        format="%Y-%m-%d",
        validators=[Optional()],
    )
    shipment_value = StringField("Shipment value", validators=[Optional(), Length(max=100)])
    tracking_status = SelectField(
        "Tracking status",
        choices=TRACKING_STATUSES,
        validators=[DataRequired()],
    )
    stage = SelectField("Stage", validators=[DataRequired()])
    customs_status = SelectField("Customs status", choices=CUSTOMS_STATUSES, validators=[DataRequired()])
    customs_reference = StringField("Customs reference", validators=[Optional(), Length(max=200)])
    payment_status = SelectField("Payment status", choices=PAYMENT_STATUSES, validators=[DataRequired()])
    pending_items = TextAreaField("Pending items (one per line)", validators=[Optional()])
    re_export_status = SelectField("Re-export status", choices=RE_EXPORT_STATUSES, validators=[DataRequired()])
    re_export_reference = StringField("Re-export reference", validators=[Optional(), Length(max=200)])
    container_no = StringField("Container / seal number", validators=[Optional(), Length(max=200)])
    warehouse_location = StringField("Warehouse location", validators=[Optional(), Length(max=200)])
    notes = TextAreaField("Notes", validators=[Optional()])
    submit = SubmitField("Save shipment")


class CreateUserForm(FlaskForm):
    username = StringField("Username", validators=[DataRequired(), Length(min=3, max=80)])
    password = PasswordField("Initial password (minimum 12 characters)", validators=[DataRequired(), Length(min=12)])
    can_view = BooleanField("View records", default=True)
    can_add = BooleanField("Add records")
    can_edit = BooleanField("Edit records")
    can_delete = BooleanField("Delete records")
    submit = SubmitField("Create user")


class ManageUserForm(FlaskForm):
    can_view = BooleanField("View records")
    can_add = BooleanField("Add records")
    can_edit = BooleanField("Edit records")
    can_delete = BooleanField("Delete records")
    new_password = PasswordField(
        "Set/reset password (leave blank to keep current)",
        validators=[Optional(), Length(min=12)],
    )
    submit = SubmitField("Save user")


class ChangePasswordForm(FlaskForm):
    current_password = PasswordField("Current password", validators=[DataRequired()])
    new_password = PasswordField(
        "New password (minimum 12 characters)",
        validators=[DataRequired(), Length(min=12)],
    )
    confirm_password = PasswordField(
        "Confirm new password",
        validators=[DataRequired(), EqualTo("new_password", message="Passwords do not match.")],
    )
    submit = SubmitField("Change password")


class EventForm(FlaskForm):
    exhibition_name = StringField(
        "Exhibition name",
        validators=[DataRequired(), Length(max=200)],
    )
    master_file_ref = StringField(
        "Master file reference",
        validators=[DataRequired(), Length(max=50)],
    )
    start_date = DateField("Exhibition start date", format="%Y-%m-%d", validators=[DataRequired()])
    end_date = DateField("Exhibition end date", format="%Y-%m-%d", validators=[DataRequired()])
    venue = StringField("Venue", validators=[DataRequired(), Length(max=200)])
    organiser_name = StringField("Organiser name", validators=[DataRequired(), Length(max=200)])
    move_in_start = DateField("Move-in start date", format="%Y-%m-%d", validators=[Optional()])
    move_in_end = DateField("Move-in end date", format="%Y-%m-%d", validators=[Optional()])
    move_out_start = DateField("Move-out start date", format="%Y-%m-%d", validators=[Optional()])
    move_out_end = DateField("Move-out end date", format="%Y-%m-%d", validators=[Optional()])
    submit = SubmitField("Save event")


class EventShipmentFileForm(FlaskForm):
    file_date = DateField("Date", format="%Y-%m-%d", default=date.today, validators=[Optional()])
    event_id = SelectField("Event name", choices=[], validators=[DataRequired()])
    transport_type = SelectField(
        "Type of transport",
        choices=[SelectChoice("", "Select transport")]
        + [SelectChoice(value, value) for value in TRANSPORT_TYPES],
        validators=[DataRequired()],
    )
    custom_transport = StringField(
        "Other transport details",
        validators=[Optional(), Length(max=100)],
    )
    on_site_handling = SelectField(
        "On-site handling",
        choices=[SelectChoice("", "Select an option")]
        + [SelectChoice(value, value) for value in ON_SITE_HANDLING_CHOICES],
        validators=[DataRequired()],
    )
    mawb_no = StringField("AWB No.", validators=[Optional(), Length(max=50)])
    hawb_no = StringField("HAWB No.", validators=[Optional(), Length(max=50)])
    hawb_departure = StringField("Departure", validators=[Optional(), Length(max=100)])
    hawb_arrival = StringField("Arrival", validators=[Optional(), Length(max=100)])
    hawb_departure_date = DateField(
        "Departure date",
        format="%Y-%m-%d",
        validators=[Optional()],
    )
    hawb_arrival_date = DateField(
        "Arrival date",
        format="%Y-%m-%d",
        validators=[Optional()],
    )
    tracking_status = SelectField(
        "Tracking status",
        choices=[SelectChoice("", "Select tracking status")]
        + [SelectChoice(value, value) for value in TRACKING_STATUSES],
        validators=[Optional()],
        default="In Transit",
    )
    customs_status = SelectField(
        "Customs status",
        choices=[SelectChoice("", "Select customs status")]
        + [SelectChoice(value, value) for value in AIR_CUSTOMS_STATUSES],
        validators=[Optional()],
        default="Not Filed",
    )
    payment_status = SelectField(
        "Payment status",
        choices=[SelectChoice("", "Select payment status")]
        + [SelectChoice(value, value) for value in PAYMENT_STATUSES],
        validators=[Optional()],
        default="Pending",
    )
    stage = SelectField(
        "Stage",
        choices=[SelectChoice("", "Select stage")]
        + [SelectChoice(value, value) for value in SHIPMENT_STAGES],
        validators=[Optional()],
        default="Received",
    )
    exhibitor_name = StringField("Exhibitor / company name", validators=[DataRequired(), Length(max=200)])
    exhibitor_contact_person = StringField("Exhibitor contact person", validators=[Optional(), Length(max=200)])
    exhibitor_email = StringField("Exhibitor email", validators=[Optional(), Length(max=254)])
    exhibitor_phone = StringField("Exhibitor phone", validators=[Optional(), Length(max=50)])
    agent_name = StringField("Agent / company name", validators=[DataRequired(), Length(max=200)])
    agent_contact_person = StringField("Agent contact person", validators=[Optional(), Length(max=200)])
    agent_email = StringField("Agent email", validators=[Optional(), Length(max=254)])
    agent_phone = StringField("Agent phone", validators=[Optional(), Length(max=50)])
    customer_name = StringField("Customer / company name", validators=[DataRequired(), Length(max=200)])
    customer_contact_person = StringField("Customer contact person", validators=[Optional(), Length(max=200)])
    customer_email = StringField("Customer email", validators=[Optional(), Length(max=254)])
    customer_phone = StringField("Customer phone", validators=[Optional(), Length(max=50)])
    payment_name = StringField("Payment", validators=[Optional(), Length(max=200)])
    payment_contact_person = StringField("Contact name", validators=[Optional(), Length(max=200)])
    payment_email = StringField("Email", validators=[Optional(), Length(max=254)])
    payment_phone = StringField("Telephone", validators=[Optional(), Length(max=50)])
    pickup_name = StringField("Pickup", validators=[Optional(), Length(max=200)])
    pickup_contact_person = StringField("Contact name", validators=[Optional(), Length(max=200)])
    pickup_email = StringField("Email", validators=[Optional(), Length(max=254)])
    pickup_phone = StringField("Telephone", validators=[Optional(), Length(max=50)])
    hall_number = StringField("Hall number", validators=[Optional(), Length(max=100)])
    booth_number = StringField("Booth number", validators=[Optional(), Length(max=100)])
    tax_formalities = SelectField(
        "Tax formalities",
        choices=[SelectChoice("", "Choice"), SelectChoice("Yes", "Yes"), SelectChoice("No", "No")],
        validators=[DataRequired()],
    )
    insurance = SelectField(
        "Insurance",
        choices=[SelectChoice("", "Choice"), SelectChoice("Yes", "Yes"), SelectChoice("No", "No")],
        validators=[DataRequired()],
    )
    insurance_value = DecimalField(
        "Insurance value (INR)",
        places=2,
        default=Decimal("0.00"),
        validators=[Optional(), NumberRange(min=0, max=Decimal("9999999999.99"))],
    )
    submit = SubmitField("Save file")


class AgentContactPersonFields(Form):
    contact_name = StringField("Contact person name", validators=[DataRequired(), Length(max=200)])
    phone = StringField("Contact phone number", validators=[DataRequired(), Length(max=50)])


class AgentProfileForm(FlaskForm):
    company_name = StringField("Company name", validators=[DataRequired(), Length(max=200)])
    country = StringField("Country", validators=[Optional(), Length(max=100)])
    address = TextAreaField("Address", validators=[Optional(), Length(max=4000)])
    city = StringField("City", validators=[Optional(), Length(max=150)])
    postal_code = StringField("ZIP / Postal code", validators=[Optional(), Length(max=30)])
    company_phone = StringField("Company phone number", validators=[Optional(), Length(max=50)])
    company_email = StringField("Company email", validators=[Optional(), Length(max=254)])
    vat_number = StringField("VAT number", validators=[Optional(), Length(max=100)])
    contacts = FieldList(FormField(AgentContactPersonFields), min_entries=1)
    submit = SubmitField("Save agent")


class CustomerForm(FlaskForm):
    company_name = StringField(
        "Customer / company name",
        validators=[DataRequired(), Length(max=200)],
    )
    contact_person_name = StringField(
        "Contact person",
        validators=[Optional(), Length(max=200)],
    )
    address = TextAreaField("Address", validators=[Optional(), Length(max=4000)])
    email = StringField("Email", validators=[Optional(), Length(max=254)])
    phone = StringField("Phone", validators=[Optional(), Length(max=50)])
    submit = SubmitField("Save customer")


def upgrade_event_file_proposal_fields():
    existing_columns = {
        column["name"]
        for column in inspect(db.engine).get_columns(EventFileProposal.__tablename__)
    }
    new_columns = {
        "recipient_contact": "VARCHAR(200) NOT NULL DEFAULT ''",
        "recipient_phone": "VARCHAR(50) NOT NULL DEFAULT ''",
        "status": "VARCHAR(30) NOT NULL DEFAULT 'DRAFT'",
        "updated_at": "DATETIME",
        "discount_type": "VARCHAR(10) NOT NULL DEFAULT 'percent'",
        "discount_value": "NUMERIC(16, 2) NOT NULL DEFAULT 0",
        "bank_account_slot": "INTEGER",
        "bank_label": "VARCHAR(80) NOT NULL DEFAULT ''",
        "bank_beneficiary": "VARCHAR(200) NOT NULL DEFAULT ''",
        "bank_name": "VARCHAR(200) NOT NULL DEFAULT ''",
        "bank_account_number": "VARCHAR(100) NOT NULL DEFAULT ''",
        "bank_account_type": "VARCHAR(40) NOT NULL DEFAULT ''",
        "bank_currency": "VARCHAR(3) NOT NULL DEFAULT ''",
        "bank_branch": "VARCHAR(200) NOT NULL DEFAULT ''",
        "bank_ifsc": "VARCHAR(30) NOT NULL DEFAULT ''",
        "bank_swift": "VARCHAR(20) NOT NULL DEFAULT ''",
    }
    for column, definition in new_columns.items():
        if column not in existing_columns:
            db.session.execute(
                text(
                    f'ALTER TABLE "{EventFileProposal.__tablename__}" '
                    f'ADD COLUMN "{column}" {definition}'
                )
            )
    db.session.execute(
        text(
            f'UPDATE "{EventFileProposal.__tablename__}" '
            'SET "discount_value" = "discount_percent" '
            'WHERE "discount_type" = \'percent\''
        )
    )
    db.session.commit()


def upgrade_event_file_delivery_note_history():
    """Backfill delivery note history from existing saved delivery notes."""
    if not inspect(db.engine).has_table(EventFileDeliveryNoteHistory.__tablename__):
        return
    existing_file_ids = {
        row.file_id
        for row in db.session.query(EventFileDeliveryNoteHistory.file_id).distinct()
    }
    for note in EventFileDeliveryNote.query.all():
        if note.file_id in existing_file_ids:
            continue
        db.session.add(
            EventFileDeliveryNoteHistory(
                file_id=note.file_id,
                delivery_date=note.delivery_date,
                condition_at_delivery=note.condition_at_delivery,
                receiver_name=note.receiver_name,
                receiver_designation=note.receiver_designation,
                receiver_phone=note.receiver_phone,
                receiver_email=note.receiver_email,
                supervisor_name=note.supervisor_name,
                supervisor_designation=note.supervisor_designation,
                supervisor_phone=note.supervisor_phone,
                created_by_user_id=note.created_by_user_id,
                generated_at=note.updated_at or note.created_at,
            )
        )
    db.session.commit()


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


def parse_legacy_date(value, field_name):
    if not value:
        return None
    for date_format in ("%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(value), date_format).date()
        except ValueError:
            continue
    raise ValueError(f"Unsupported {field_name} in shipments.json: {value!r}")


def utc_now():
    return datetime.now(timezone.utc)


def as_utc(value):
    if value.tzinfo is None:
        return value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc)


def upgrade_user_permissions():
    existing_columns = {
        column["name"] for column in inspect(db.engine).get_columns(User.__tablename__)
    }
    new_columns = {
        "can_view": "BOOLEAN NOT NULL DEFAULT TRUE",
        "can_add": "BOOLEAN NOT NULL DEFAULT FALSE",
        "can_edit": "BOOLEAN NOT NULL DEFAULT FALSE",
        "can_delete": "BOOLEAN NOT NULL DEFAULT FALSE",
    }
    for column, definition in new_columns.items():
        if column not in existing_columns:
            db.session.execute(
                text(f'ALTER TABLE "{User.__tablename__}" ADD COLUMN "{column}" {definition}')
            )
    shipment_columns = {
        column["name"] for column in inspect(db.engine).get_columns(Shipment.__tablename__)
    }
    if "tracking_status" not in shipment_columns:
        db.session.execute(
            text(
                f'ALTER TABLE "{Shipment.__tablename__}" '
                'ADD COLUMN "tracking_status" VARCHAR(40) NOT NULL DEFAULT \'In Transit\''
            )
        )
        db.session.execute(
            text(
                f'UPDATE "{Shipment.__tablename__}" '
                'SET "tracking_status" = "stage" '
                'WHERE "stage" IN (\'In Transit\', \'Out for Delivery\', \'Delivered\', \'Delayed\')'
            )
        )
    db.session.commit()


def upgrade_event_file_air_details():
    existing_columns = {
        column["name"]
        for column in inspect(db.engine).get_columns(
            EventShipmentFileAirDetails.__tablename__
        )
    }
    new_columns = {
        "hawb_departure_date": "DATE",
        "tracking_status": "VARCHAR(40) NOT NULL DEFAULT 'In Transit'",
        "customs_status": "VARCHAR(100) NOT NULL DEFAULT 'Not Filed'",
        "payment_status": "VARCHAR(40) NOT NULL DEFAULT 'Pending'",
        "stage": "VARCHAR(100) NOT NULL DEFAULT 'Received'",
    }
    for column, definition in new_columns.items():
        if column not in existing_columns:
            db.session.execute(
                text(
                    f'ALTER TABLE "{EventShipmentFileAirDetails.__tablename__}" '
                    f'ADD COLUMN "{column}" {definition}'
                )
            )
    db.session.commit()


def upgrade_agent_profile_location_fields():
    existing_columns = {
        column["name"] for column in inspect(db.engine).get_columns(AgentProfile.__tablename__)
    }
    new_columns = {
        "city": "VARCHAR(150) NOT NULL DEFAULT ''",
        "postal_code": "VARCHAR(30) NOT NULL DEFAULT ''",
    }
    for column, definition in new_columns.items():
        if column not in existing_columns:
            db.session.execute(
                text(
                    f'ALTER TABLE "{AgentProfile.__tablename__}" '
                    f'ADD COLUMN "{column}" {definition}'
                )
            )
    db.session.commit()


def upgrade_contact_address():
    existing_columns = {
        column["name"] for column in inspect(db.engine).get_columns(Contact.__tablename__)
    }
    if "address" not in existing_columns:
        db.session.execute(
            text(
                f'ALTER TABLE "{Contact.__tablename__}" '
                'ADD COLUMN "address" TEXT NOT NULL DEFAULT \'\''
            )
        )
        db.session.commit()


def audit_event(action, entity_type, entity_id, entity_label, details):
    actor = current_user if current_user.is_authenticated else None
    db.session.add(
        AuditLog(
            actor_user_id=actor.id if actor else None,
            actor_username=actor.username if actor else "system",
            action=action,
            entity_type=entity_type,
            entity_id=entity_id,
            entity_label=entity_label or "",
            details_json=json.dumps(details, ensure_ascii=False, sort_keys=True),
            created_at=utc_now(),
        )
    )


def shipment_snapshot(shipment):
    fields = (
        "reference",
        "agent_name",
        "client",
        "exhibitor_name",
        "exhibition_name",
        "supplier",
        "origin",
        "destination",
        "tracking_status",
        "arrival_date",
        "arrival_time",
        "expected_reexport_date",
        "shipment_value",
        "stage",
        "customs_status",
        "customs_reference",
        "payment_status",
        "pending_items",
        "re_export_status",
        "re_export_reference",
        "container_no",
        "warehouse_location",
        "notes",
    )
    snapshot = {}
    for field in fields:
        value = getattr(shipment, field)
        snapshot[field] = value.isoformat() if hasattr(value, "isoformat") else value
    return snapshot


def event_snapshot(event):
    fields = (
        "department",
        "exhibition_name",
        "master_file_ref",
        "start_date",
        "end_date",
        "venue",
        "organiser_name",
        "move_in_start",
        "move_in_end",
        "move_out_start",
        "move_out_end",
    )
    snapshot = {}
    for field in fields:
        value = getattr(event, field)
        snapshot[field] = value.isoformat() if isinstance(value, date) else value
    return snapshot


def contact_snapshot(contact):
    return {
        "kind": contact.kind,
        "company_name": contact.company_name,
        "contact_person_name": contact.contact_person_name,
        "address": contact.address,
        "email": contact.email,
        "phone": contact.phone,
    }


def event_file_snapshot(file_record):
    details = file_record.details
    snapshot = {
        "file_reference": file_record.file_reference,
        "event_id": file_record.event_id,
        "file_date": (details.file_date if details else file_record.file_date).isoformat(),
        "transport_type": file_record.transport_type,
        "custom_transport": file_record.custom_transport,
        "on_site_handling": file_record.on_site_handling,
        "hall_number": details.hall_number if details else "",
        "booth_number": details.booth_number if details else "",
        "tax_formalities": details.tax_formalities if details else "",
        "insurance": details.insurance if details else "",
        "insurance_value": str(details.insurance_value) if details else "0.00",
    }
    for kind in EVENT_FILE_CONTACT_KINDS:
        contact = getattr(file_record, f"{kind}_contact")
        snapshot[f"{kind}_contact"] = contact_snapshot(contact) if contact else None
    for kind in ("payment", "pickup"):
        contact = getattr(details, f"{kind}_contact") if details else None
        snapshot[f"{kind}_contact"] = contact_snapshot(contact) if contact else None
    air_details = file_record.air_details
    snapshot["air_shipment"] = (
        {
            "mawb_no": air_details.mawb_no,
            "hawb_no": air_details.hawb_no,
            "hawb_departure": air_details.hawb_departure,
            "hawb_arrival": air_details.hawb_arrival,
            "hawb_departure_date": (
                air_details.hawb_departure_date.isoformat()
                if air_details.hawb_departure_date
                else None
            ),
            "hawb_arrival_date": (
                air_details.hawb_arrival_date.isoformat()
                if air_details.hawb_arrival_date
                else None
            ),
            "tracking_status": air_details.tracking_status,
            "customs_status": air_details.customs_status,
            "payment_status": air_details.payment_status,
            "stage": air_details.stage,
        }
        if air_details
        else None
    )
    snapshot["goods_items"] = [
        {
            "direction": item.direction,
            "marking": item.marking,
            "pieces": item.pieces,
            "length_cm": str(item.length_cm),
            "width_cm": str(item.width_cm),
            "height_cm": str(item.height_cm),
            "pack_type": item.pack_type,
            "pack_other": item.pack_other,
            "weight_kg": str(item.weight_kg),
            "description": item.description,
            "clearance_method": item.clearance_method,
        }
        for item in file_record.goods_items
    ]
    return snapshot


def event_goods_form_data_present():
    return request.form.get("goods-form-present") == "1"


def empty_event_goods_row():
    return {
        "marking": "",
        "pieces": "1",
        "length_cm": "",
        "width_cm": "",
        "height_cm": "",
        "pack_type": "Wooden Packing",
        "pack_other": "",
        "weight_kg": "",
        "description": "",
        "clearance_method": "ATA",
    }


def event_goods_form_rows(file_record=None):
    if request.method == "POST" and event_goods_form_data_present():
        rows_by_direction = {direction: {} for direction in GOODS_DIRECTIONS}
        for direction in GOODS_DIRECTIONS:
            prefix = f"{direction}-goods-"
            for key in request.form.keys():
                if not key.startswith(prefix):
                    continue
                row_index, separator, field_name = key[len(prefix):].partition("-")
                if (
                    not separator
                    or not row_index.isdigit()
                    or field_name not in GOODS_FORM_FIELDS
                ):
                    continue
                rows_by_direction[direction].setdefault(int(row_index), {})[field_name] = request.form.get(key, "")
        return {
            direction: [
                {**empty_event_goods_row(), **row_values}
                for _, row_values in sorted(rows_by_direction[direction].items())
            ] or [empty_event_goods_row()]
            for direction in GOODS_DIRECTIONS
        }

    if file_record is not None:
        rows_by_direction = {
            direction: [
                {
                    "marking": item.marking,
                    "pieces": str(item.pieces),
                    "length_cm": str(item.length_cm),
                    "width_cm": str(item.width_cm),
                    "height_cm": str(item.height_cm),
                    "pack_type": item.pack_type,
                    "pack_other": item.pack_other,
                    "weight_kg": str(item.weight_kg),
                    "description": item.description,
                    "clearance_method": item.clearance_method,
                }
                for item in file_record.goods_items
                if item.direction == direction
            ]
            for direction in GOODS_DIRECTIONS
        }
        return {
            direction: rows_by_direction[direction] or [empty_event_goods_row()]
            for direction in GOODS_DIRECTIONS
        }

    return {direction: [empty_event_goods_row()] for direction in GOODS_DIRECTIONS}


def validate_event_goods_form_rows(rows_by_direction):
    parsed_rows = {direction: [] for direction in GOODS_DIRECTIONS}
    errors = []
    max_amount = Decimal("9999999999.99")

    for direction in GOODS_DIRECTIONS:
        for index, row in enumerate(rows_by_direction[direction], start=1):
            values = {
                field: str(row.get(field, "") or "").strip()
                for field in GOODS_FORM_FIELDS
            }
            if not any(
                values[field]
                for field in (
                    "marking",
                    "length_cm",
                    "width_cm",
                    "height_cm",
                    "pack_other",
                    "weight_kg",
                    "description",
                )
            ):
                continue

            row_label = f"{direction.title()} row {index}"
            row_errors = []
            if not values["marking"]:
                row_errors.append(f"{row_label}: Goods marking is required.")
            elif len(values["marking"]) > 120:
                row_errors.append(f"{row_label}: Goods marking must be 120 characters or fewer.")

            try:
                pieces = int(values["pieces"])
            except (TypeError, ValueError):
                pieces = 0
            if pieces < 1 or pieces > 1_000_000:
                row_errors.append(f"{row_label}: Pieces must be a whole number between 1 and 1,000,000.")

            amounts = {}
            for field, label in (
                ("length_cm", "Length"),
                ("width_cm", "Width"),
                ("height_cm", "Height"),
                ("weight_kg", "Weight"),
            ):
                raw_amount = values[field]
                try:
                    amount = Decimal(raw_amount) if raw_amount else None
                except InvalidOperation:
                    amount = None
                if amount is None or not amount.is_finite():
                    row_errors.append(f"{row_label}: Enter a valid {label.lower()}.")
                elif amount < 0 or amount > max_amount or (field != "weight_kg" and amount == 0):
                    row_errors.append(f"{row_label}: {label} is outside the allowed range.")
                else:
                    amounts[field] = amount

            if values["pack_type"] not in PACKING_TYPES:
                row_errors.append(f"{row_label}: Select a valid packing type.")
            elif values["pack_type"] == "Other" and not values["pack_other"]:
                row_errors.append(f"{row_label}: Specify the other packing type.")
            if len(values["pack_other"]) > 120:
                row_errors.append(f"{row_label}: Other packing type must be 120 characters or fewer.")

            if len(values["description"]) > 500:
                row_errors.append(f"{row_label}: Description must be 500 characters or fewer.")
            if values["clearance_method"] not in CLEARANCE_METHODS:
                row_errors.append(f"{row_label}: Select a valid clearance method.")

            errors.extend(row_errors)
            if row_errors:
                continue
            parsed_rows[direction].append(
                {
                    "marking": values["marking"],
                    "pieces": pieces,
                    "length_cm": amounts["length_cm"],
                    "width_cm": amounts["width_cm"],
                    "height_cm": amounts["height_cm"],
                    "pack_type": values["pack_type"],
                    "pack_other": values["pack_other"],
                    "weight_kg": amounts["weight_kg"],
                    "description": values["description"],
                    "clearance_method": values["clearance_method"],
                }
            )
    return parsed_rows, errors


def save_event_goods(file_record, parsed_rows):
    current_rows = {direction: [] for direction in GOODS_DIRECTIONS}
    for item in file_record.goods_items:
        current_rows[item.direction].append(
            {
                "marking": item.marking,
                "pieces": item.pieces,
                "length_cm": item.length_cm,
                "width_cm": item.width_cm,
                "height_cm": item.height_cm,
                "pack_type": item.pack_type,
                "pack_other": item.pack_other,
                "weight_kg": item.weight_kg,
                "description": item.description,
                "clearance_method": item.clearance_method,
            }
        )
    if current_rows == parsed_rows:
        return

    file_record.goods_items.clear()
    for direction in GOODS_DIRECTIONS:
        for sort_order, values in enumerate(parsed_rows[direction]):
            file_record.goods_items.append(
                EventShipmentGood(
                    direction=direction,
                    sort_order=sort_order,
                    **values,
                )
            )


def calculate_event_goods_metrics(goods_row, transport_type):
    dimensions = (
        Decimal(str(goods_row["length_cm"]))
        * Decimal(str(goods_row["width_cm"]))
        * Decimal(str(goods_row["height_cm"]))
        * Decimal(str(goods_row["pieces"]))
    )
    cbm = (dimensions / Decimal("1000000")).quantize(
        Decimal("0.01"),
        rounding=ROUND_HALF_UP,
    )
    if transport_type == "Air":
        volumetric_weight = (dimensions / Decimal("6000")).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    elif transport_type == "Sea":
        volumetric_weight = (dimensions / Decimal("1000")).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    else:
        volumetric_weight = None
    return {"cbm": cbm, "volumetric_weight_kg": volumetric_weight}


def saved_contacts_by_kind():
    return {
        kind: [
            {
                "company_name": contact.company_name,
                "contact_person_name": contact.contact_person_name,
                "email": contact.email,
                "phone": contact.phone,
            }
            for contact in Contact.query.filter_by(kind=kind).order_by(Contact.company_name.asc()).all()
        ]
        for kind in CONTACT_KINDS
    }


def upsert_contact(kind, company_name, contact_person_name, email, phone):
    company_name = company_name.strip()
    if not company_name:
        return None

    name_key = company_name.casefold()
    contact = Contact.query.filter_by(kind=kind, name_key=name_key).first()
    values = {
        "company_name": company_name,
        "contact_person_name": contact_person_name.strip(),
        "email": email.strip(),
        "phone": phone.strip(),
    }
    if contact is None:
        contact = Contact(
            kind=kind,
            name_key=name_key,
            created_by_user_id=current_user.id,
            **values,
        )
        db.session.add(contact)
        db.session.flush()
        audit_event("create", "contact", contact.id, company_name, {"fields": contact_snapshot(contact)})
        return contact

    before = contact_snapshot(contact)
    contact.company_name = company_name
    for field in ("contact_person_name", "email", "phone"):
        setattr(contact, field, values[field])
    after = contact_snapshot(contact)
    if before != after:
        contact.updated_at = utc_now()
        audit_event(
            "edit",
            "contact",
            contact.id,
            contact.company_name,
            {
                "fields": {
                    field: {"before": before[field], "after": after[field]}
                    for field in before
                    if before[field] != after[field]
                }
            },
        )
    return contact


def event_file_event_options(events):
    options = []
    for event in events:
        date_range = (
            f"{event.start_date.strftime('%d-%m-%Y')} - "
            f"{event.end_date.strftime('%d-%m-%Y')}"
        )
        reference = event.master_file_ref or "Master ref required"
        options.append(
            {
                "id": str(event.id),
                "name": event.exhibition_name,
                "reference": reference,
                "dates": date_range,
                "label": f"{event.exhibition_name} - {reference} ({date_range})",
            }
        )
    return options


def make_event_file_form(events, file_record=None, selected_event_id=None):
    form = EventShipmentFileForm(obj=file_record)
    event_options = event_file_event_options(events)
    form.event_id.choices = [SelectChoice("", "Choose a saved event")] + [
        SelectChoice(option["id"], option["label"]) for option in event_options
    ]
    if file_record is not None:
        details = file_record.details
        form.file_date.data = details.file_date if details else file_record.file_date
    if file_record is not None and request.method == "GET":
        form.event_id.data = str(file_record.event_id)
        if file_record.transport_type == "Custom":
            form.transport_type.data = "Other"
        for field in ("hall_number", "booth_number", "tax_formalities", "insurance", "insurance_value"):
            if details:
                getattr(form, field).data = getattr(details, field)
        if file_record.air_details:
            for field in AIR_SHIPMENT_FORM_FIELDS:
                getattr(form, field).data = getattr(file_record.air_details, field)
        for kind in EVENT_FILE_CONTACT_KINDS:
            contact = getattr(file_record, f"{kind}_contact")
            if contact:
                getattr(form, f"{kind}_name").data = contact.company_name
                getattr(form, f"{kind}_contact_person").data = contact.contact_person_name
                getattr(form, f"{kind}_email").data = contact.email
                getattr(form, f"{kind}_phone").data = contact.phone
        if details:
            for kind in ("payment", "pickup"):
                contact = getattr(details, f"{kind}_contact")
                if contact:
                    getattr(form, f"{kind}_name").data = contact.company_name
                    getattr(form, f"{kind}_contact_person").data = contact.contact_person_name
                    getattr(form, f"{kind}_email").data = contact.email
                    getattr(form, f"{kind}_phone").data = contact.phone
    elif selected_event_id is not None and request.method == "GET":
        form.event_id.data = str(selected_event_id)
    return form


def apply_event_file_form(form, file_record):
    selected_event = db.session.get(Event, int(form.event_id.data))
    if selected_event is None:
        abort(404)
    file_record.event = selected_event
    file_record.event_id = selected_event.id
    file_record.transport_type = form.transport_type.data
    file_record.custom_transport = (
        (form.custom_transport.data or "").strip()
        if form.transport_type.data == "Other"
        else ""
    )
    file_record.on_site_handling = form.on_site_handling.data
    if form.transport_type.data != "Air":
        file_record.air_details = None
    elif any(field in request.form for field in AIR_SHIPMENT_FORM_FIELDS):
        existing_air_details = file_record.air_details
        air_values = {}
        for field in AIR_SHIPMENT_FORM_FIELDS:
            if field in request.form:
                value = getattr(form, field).data
                air_values[field] = value.strip() if isinstance(value, str) else value
            elif existing_air_details:
                air_values[field] = getattr(existing_air_details, field)
        if any(value for value in air_values.values()):
            air_details = existing_air_details or EventShipmentFileAirDetails(file=file_record)
            for field, value in air_values.items():
                setattr(air_details, field, value)
            file_record.air_details = air_details
        else:
            file_record.air_details = None
    for kind in EVENT_FILE_CONTACT_KINDS:
        contact = upsert_contact(
            kind,
            getattr(form, f"{kind}_name").data,
            getattr(form, f"{kind}_contact_person").data or "",
            getattr(form, f"{kind}_email").data or "",
            getattr(form, f"{kind}_phone").data or "",
        )
        setattr(file_record, f"{kind}_contact", contact)

    details = file_record.details or EventShipmentFileDetails(file=file_record)
    if file_record.id is None:
        details.file_date = date.today()
    elif file_record.details:
        details.file_date = file_record.details.file_date
    else:
        details.file_date = file_record.created_at.date()
    details.hall_number = (form.hall_number.data or "").strip()
    details.booth_number = (form.booth_number.data or "").strip()
    details.tax_formalities = form.tax_formalities.data
    details.insurance = form.insurance.data
    details.insurance_value = form.insurance_value.data or Decimal("0.00")
    for kind in ("payment", "pickup"):
        contact = upsert_contact(
            kind,
            getattr(form, f"{kind}_name").data or "",
            getattr(form, f"{kind}_contact_person").data or "",
            getattr(form, f"{kind}_email").data or "",
            getattr(form, f"{kind}_phone").data or "",
        )
        setattr(details, f"{kind}_contact", contact)
    file_record.details = details


def agent_profile_snapshot(profile):
    return {
        "country": profile.country,
        "address": profile.address,
        "city": profile.city,
        "postal_code": profile.postal_code,
        "company_phone": profile.company_phone,
        "company_email": profile.company_email,
        "vat_number": profile.vat_number,
        "contacts": [
            {"contact_name": person.contact_name, "phone": person.phone}
            for person in profile.contact_people
        ],
    }


def populate_agent_profile_form(form, contact):
    profile = contact.agent_profile
    form.company_name.data = contact.company_name
    form.country.data = profile.country if profile else ""
    form.address.data = profile.address if profile else ""
    form.city.data = profile.city if profile else ""
    form.postal_code.data = profile.postal_code if profile else ""
    form.company_phone.data = profile.company_phone if profile else contact.phone
    form.company_email.data = profile.company_email if profile else contact.email
    form.vat_number.data = profile.vat_number if profile else ""

    people = (
        [
            {"contact_name": person.contact_name, "phone": person.phone}
            for person in profile.contact_people
        ]
        if profile and profile.contact_people
        else [{"contact_name": contact.contact_person_name, "phone": contact.phone}]
    )
    for index, person in enumerate(people):
        entry = form.contacts.entries[0] if index == 0 else form.contacts.append_entry()
        entry.form.contact_name.data = person["contact_name"]
        entry.form.phone.data = person["phone"]


def save_agent_profile(form, contact=None):
    people = [
        {
            "contact_name": entry.form.contact_name.data.strip(),
            "phone": entry.form.phone.data.strip(),
        }
        for entry in form.contacts.entries
    ]
    first_person = people[0]
    if contact is None:
        contact = upsert_contact(
            "agent",
            form.company_name.data,
            first_person["contact_name"],
            form.company_email.data or "",
            form.company_phone.data or "",
        )
    else:
        before_contact = contact_snapshot(contact)
        contact.company_name = form.company_name.data.strip()
        contact.name_key = contact.company_name.casefold()
        contact.contact_person_name = first_person["contact_name"]
        contact.email = (form.company_email.data or "").strip()
        contact.phone = (form.company_phone.data or "").strip()
        after_contact = contact_snapshot(contact)
        if before_contact != after_contact:
            audit_event(
                "edit",
                "contact",
                contact.id,
                contact.company_name,
                {
                    "fields": {
                        key: {"before": before_contact[key], "after": after_contact[key]}
                        for key in before_contact
                        if before_contact[key] != after_contact[key]
                    }
                },
            )
            contact.updated_at = utc_now()
    profile = contact.agent_profile
    is_new_profile = profile is None
    before = agent_profile_snapshot(profile) if profile else None
    if profile is None:
        profile = AgentProfile(contact=contact)
        db.session.add(profile)

    profile.country = (form.country.data or "").strip()
    profile.address = (form.address.data or "").strip()
    profile.city = (form.city.data or "").strip()
    profile.postal_code = (form.postal_code.data or "").strip()
    profile.company_phone = (form.company_phone.data or "").strip()
    profile.company_email = (form.company_email.data or "").strip()
    profile.vat_number = (form.vat_number.data or "").strip()
    profile.contact_people = [
        AgentContactPerson(**person)
        for person in people
    ]
    db.session.flush()
    after = agent_profile_snapshot(profile)

    if is_new_profile:
        audit_event("create", "agent_profile", profile.id, contact.company_name, {"fields": after})
        contact.updated_at = utc_now()
    elif before != after:
        audit_event(
            "edit",
            "agent_profile",
            profile.id,
            contact.company_name,
            {
                "fields": {
                    key: {"before": before[key], "after": after[key]}
                    for key in before
                    if before[key] != after[key]
                }
            },
        )
        contact.updated_at = utc_now()
    return contact


def get_department(department_slug):
    for department in DEPARTMENTS:
        if department["slug"] == department_slug:
            return department
    abort(404)


def optional_query_integer(name):
    value = request.args.get(name, "").strip()
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        abort(400)


def event_date_ranges_valid(form):
    pairs = (
        (form.start_date, form.end_date, "Exhibition"),
        (form.move_in_start, form.move_in_end, "Move-in"),
        (form.move_out_start, form.move_out_end, "Move-out"),
    )
    valid = True
    for start_field, end_field, label in pairs:
        start_date = start_field.data
        end_date = end_field.data
        if bool(start_date) != bool(end_date):
            missing_field = end_field if start_date else start_field
            missing_field.errors.append(f"Please enter both the {label.lower()} start and end dates.")
            valid = False
        elif start_date and end_date and end_date < start_date:
            end_field.errors.append("The end date must be on or after the start date.")
            valid = False
    return valid


def event_master_reference_valid(form, event=None):
    master_ref = form.master_file_ref.data.strip()
    existing = EventMasterReference.query.filter_by(master_ref_key=master_ref.casefold()).first()
    if existing and (event is None or existing.event_id != event.id):
        form.master_file_ref.errors.append("This master file reference is already assigned to another event.")
        return False
    if (
        event
        and event.master_file_ref != master_ref
        and EventShipmentFileReference.query.filter_by(event_id=event.id).first()
    ):
        form.master_file_ref.errors.append("The master file reference cannot change after file references are issued.")
        return False
    return True


def ensure_event_file_sequences(event_id):
    existing_codes = {
        row.transport_code
        for row in EventFileSequence.query.filter_by(event_id=event_id).all()
    }
    db.session.add_all(
        EventFileSequence(event_id=event_id, transport_code=code, next_sequence=1)
        for code in FILE_REFERENCE_CODES
        if code not in existing_codes
    )
    db.session.flush()


def allocate_event_file_sequence(event_id, transport_code):
    ensure_event_file_sequences(event_id)
    counter_filter = (
        EventFileSequence.event_id == event_id,
        EventFileSequence.transport_code == transport_code,
    )
    db.session.execute(
        update(EventFileSequence)
        .where(*counter_filter)
        .values(next_sequence=EventFileSequence.next_sequence + 1)
    )
    next_sequence = db.session.execute(
        select(EventFileSequence.next_sequence).where(*counter_filter)
    ).scalar_one()
    return next_sequence - 1


def file_transport_reference_code(transport_type, other_transport):
    if transport_type == "Other" and (other_transport or "").strip().casefold() == "courier":
        return "C"
    return TRANSPORT_REFERENCE_CODES[transport_type]


def assign_event_file_reference(file_record, transport_code):
    event = file_record.event
    if not event.master_file_ref:
        raise ValueError(f"Event {event.id} does not have a master file reference.")
    sequence = allocate_event_file_sequence(event.id, transport_code)
    reference = f"{event.master_file_ref}/{transport_code}{sequence:03d}"
    if file_record.reference_record is None:
        file_record.reference_record = EventShipmentFileReference(
            event_id=event.id,
            transport_code=transport_code,
            sequence=sequence,
            reference=reference,
        )
    else:
        file_record.reference_record.event_id = event.id
        file_record.reference_record.transport_code = transport_code
        file_record.reference_record.sequence = sequence
        file_record.reference_record.reference = reference


def event_file_form_valid(form):
    if not form.validate_on_submit():
        return False
    if form.transport_type.data == "Other" and not (form.custom_transport.data or "").strip():
        form.custom_transport.errors.append("Please specify the other transport type.")
        return False
    selected_event = db.session.get(Event, int(form.event_id.data))
    if selected_event and not selected_event.master_file_ref:
        form.event_id.errors.append("Add a master file reference to this event before creating files.")
        return False
    return True


def export_value(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return as_utc(value).astimezone(DISPLAY_TIMEZONE).strftime("%d-%m-%Y %I:%M:%S %p")
    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    return str(value)


def filtered_shipment_query(search, status_filter):
    query = Shipment.query
    if search:
        pattern = f"%{search}%"
        query = query.filter(
            or_(
                Shipment.reference.ilike(pattern),
                Shipment.client.ilike(pattern),
                Shipment.supplier.ilike(pattern),
                Shipment.origin.ilike(pattern),
                Shipment.destination.ilike(pattern),
            )
        )
    if status_filter:
        query = query.filter_by(tracking_status=status_filter)
    return query.order_by(Shipment.arrival_date.asc(), Shipment.reference.asc())


def filtered_event_query(department, search):
    query = Event.query.filter_by(department=department)
    if search:
        pattern = f"%{search}%"
        query = query.outerjoin(
            EventMasterReference,
            EventMasterReference.event_id == Event.id,
        )
        query = query.filter(
            or_(
                Event.exhibition_name.ilike(pattern),
                Event.venue.ilike(pattern),
                Event.organiser_name.ilike(pattern),
                EventMasterReference.master_ref.ilike(pattern),
            )
        )
    return query.order_by(Event.start_date.asc(), Event.exhibition_name.asc())


EVENT_FILE_SEARCH_FIELDS = (
    "file_code",
    "agent",
    "exhibitor",
    "awb",
    "event",
    "customer",
)


def filtered_event_file_query(department, search, event_id=None, search_field="all"):
    query = EventShipmentFile.query.join(Event).filter(Event.department == department)
    if event_id is not None:
        query = query.filter(EventShipmentFile.event_id == event_id)
    if search:
        # Case-insensitive "contains" match, regardless of upper/lower case input.
        pattern = f"%{search}%"
        AgentContact = db.aliased(Contact)
        ExhibitorContact = db.aliased(Contact)
        CustomerContact = db.aliased(Contact)
        query = query.outerjoin(AgentContact, EventShipmentFile.agent_contact_id == AgentContact.id)
        query = query.outerjoin(
            ExhibitorContact, EventShipmentFile.exhibitor_contact_id == ExhibitorContact.id
        )
        query = query.outerjoin(
            CustomerContact, EventShipmentFile.customer_contact_id == CustomerContact.id
        ).outerjoin(
            EventShipmentFileAirDetails,
            EventShipmentFileAirDetails.file_id == EventShipmentFile.id,
        ).outerjoin(
            EventShipmentFileReference,
            EventShipmentFileReference.file_id == EventShipmentFile.id,
        ).outerjoin(
            EventMasterReference,
            EventMasterReference.event_id == Event.id,
        )
        file_code_pattern = f"%{search}%"

        field_conditions = {
            "file_code": [
                EventShipmentFileReference.reference.ilike(file_code_pattern),
                EventShipmentFile.id.cast(db.String).ilike(file_code_pattern),
            ],
            "agent": [AgentContact.company_name.ilike(pattern)],
            "exhibitor": [ExhibitorContact.company_name.ilike(pattern)],
            "awb": [
                EventShipmentFileAirDetails.mawb_no.ilike(pattern),
                EventShipmentFileAirDetails.hawb_no.ilike(pattern),
            ],
            "event": [
                Event.exhibition_name.ilike(pattern),
                EventMasterReference.master_ref.ilike(pattern),
            ],
            "customer": [CustomerContact.company_name.ilike(pattern)],
            "all": [
                Event.exhibition_name.ilike(pattern),
                EventMasterReference.master_ref.ilike(pattern),
                EventShipmentFile.transport_type.ilike(pattern),
                EventShipmentFile.on_site_handling.ilike(pattern),
                EventShipmentFileReference.reference.ilike(pattern),
                EventShipmentFileAirDetails.mawb_no.ilike(pattern),
                EventShipmentFileAirDetails.hawb_no.ilike(pattern),
                AgentContact.company_name.ilike(pattern),
                ExhibitorContact.company_name.ilike(pattern),
                CustomerContact.company_name.ilike(pattern),
            ],
        }
        conditions = field_conditions.get(search_field, field_conditions["all"])
        query = query.filter(or_(*conditions))
    return query.order_by(Event.start_date.asc(), EventShipmentFile.created_at.desc())


def create_excel_export(records, fields=EXPORT_FIELDS, sheet_title="Shipments"):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_title
    worksheet.freeze_panes = "A2"
    worksheet.append([label for _, label in fields])

    header_fill = PatternFill(fill_type="solid", fgColor="102A43")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for record in records:
        row_values = [export_value(getattr(record, field)) for field, _ in fields]
        worksheet.append(row_values)
        for cell in worksheet[worksheet.max_row]:
            cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    worksheet.auto_filter.ref = worksheet.dimensions
    for column_index, (_, label) in enumerate(fields, start=1):
        values = [str(worksheet.cell(row=row, column=column_index).value or "") for row in range(1, worksheet.max_row + 1)]
        width = min(max(max((len(value) for value in values), default=len(label)) + 2, 14), 40)
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


PROPOSAL_MAX_UPLOAD_BYTES = 10 * 1024 * 1024
PROPOSAL_STATUS_OPTIONS = (
    "DRAFT",
    "SENT TO CLIENT",
    "SENT TO AGENT",
    "APPROVED",
)
PROPOSAL_COMPANY_BRANCH = "Delhi"
PROPOSAL_COMPANY_ADDRESS = (
    "1, Commercial Complex, Pocket H & J, Sarita Vihar, New Delhi - 110076"
)
PROPOSAL_COMPANY_GSTIN = "07AAACR0251M1ZH"
PROPOSAL_COMPANY_EMAIL = "rerid@rogersworldwideindia.com"
PROPOSAL_COMPANY_WEBSITE = "www.rogersworldwideindia.com"
PROPOSAL_CURRENCY_OPTIONS = sorted(
    (
        (currency.alpha_3, currency.name)
        for currency in pycountry.currencies
        if getattr(currency, "alpha_3", None)
    ),
    key=lambda option: (option[0] != "INR", option[0]),
)


def parse_proposal_decimal(value, label, default=None, minimum=None, maximum=None):
    raw_value = (value or "").strip()
    if not raw_value:
        if default is not None:
            return default
        raise ValueError(f"{label} is required.")
    try:
        parsed_value = Decimal(raw_value)
    except InvalidOperation as error:
        raise ValueError(f"{label} must be a valid number.") from error
    if not parsed_value.is_finite():
        raise ValueError(f"{label} must be a valid number.")
    if minimum is not None and parsed_value < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    if maximum is not None and parsed_value > maximum:
        raise ValueError(f"{label} cannot exceed {maximum}.")
    return parsed_value


def read_proposal_upload(upload):
    if upload is None or not upload.filename:
        return "", "", None

    filename = secure_filename(upload.filename)
    extension = Path(filename).suffix.casefold()
    data = upload.stream.read(PROPOSAL_MAX_UPLOAD_BYTES + 1)
    if len(data) > PROPOSAL_MAX_UPLOAD_BYTES:
        raise ValueError("The quotation upload must be 10 MB or smaller.")

    if extension == ".pdf":
        if not data.startswith(b"%PDF-"):
            raise ValueError("The uploaded file is not a valid PDF.")
        mimetype = "application/pdf"
    elif extension == ".docx":
        try:
            with ZipFile(BytesIO(data)) as document:
                document_names = set(document.namelist())
        except BadZipFile as error:
            raise ValueError("The uploaded Word document is not valid.") from error
        if not {"[Content_Types].xml", "word/document.xml"}.issubset(document_names):
            raise ValueError("The uploaded file is not a valid Word document.")
        mimetype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        raise ValueError("Upload a PDF or Word (.docx) quotation.")

    if not filename:
        raise ValueError("The uploaded quotation must have a valid filename.")
    return filename, mimetype, data


def build_event_file_proposal_pdf(proposal):
    output = BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=16 * mm,
        leftMargin=16 * mm,
        topMargin=15 * mm,
        bottomMargin=16 * mm,
        title=f"Quotation {proposal.quote_number}",
        author="R.E. Rogers India Pvt. Ltd.",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "QuotationTitle",
        parent=styles["Title"],
        textColor=colors.HexColor("#183541"),
        alignment=2,
        fontSize=20,
        leading=24,
    )
    small_style = ParagraphStyle(
        "QuotationSmall",
        parent=styles["BodyText"],
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#536873"),
    )
    body_style = ParagraphStyle(
        "QuotationBody",
        parent=styles["BodyText"],
        fontSize=9,
        leading=13,
        splitLongWords=True,
    )
    brand = Paragraph(
        "<b>R.E. Rogers India Pvt. Ltd.</b><br/><font color='#087da3'>"
        "EVENT &amp; LOGISTICS SERVICES</font>",
        ParagraphStyle(
            "QuotationBrand",
            parent=styles["BodyText"],
            fontSize=12,
            leading=17,
            textColor=colors.HexColor("#183541"),
        ),
    )
    logo_path = BASE_DIR / "static" / "rogers-40-years-logo.jpeg"
    brand_cell = (
        PDFImage(str(logo_path), width=46 * mm, height=18 * mm, kind="proportional")
        if logo_path.is_file()
        else brand
    )
    story = [
        Table(
            [[brand_cell, Paragraph("QUOTATION", title_style)]],
            colWidths=[88 * mm, 88 * mm],
            style=TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                    ("LINEBELOW", (0, 0), (-1, -1), 1.2, colors.HexColor("#08a5d2")),
                ]
            ),
        ),
        Spacer(1, 6 * mm),
    ]
    story.append(
        Paragraph(
            "<b>R.E. Rogers India Pvt. Ltd. · Delhi Office</b><br/>"
            f"{html.escape(PROPOSAL_COMPANY_ADDRESS)}<br/>"
            f"GSTIN: {html.escape(PROPOSAL_COMPANY_GSTIN)}<br/>"
            f"Email: {html.escape(PROPOSAL_COMPANY_EMAIL)}<br/>"
            f"Website: {html.escape(PROPOSAL_COMPANY_WEBSITE)}",
            small_style,
        )
    )
    story.append(Spacer(1, 5 * mm))

    event_file = proposal.file
    event = event_file.event
    quote_details = [
        [
            Paragraph("<b>Quotation no.</b>", body_style),
            Paragraph(html.escape(proposal.quote_number or "—"), body_style),
            Paragraph("<b>Issue date</b>", body_style),
            Paragraph(proposal.issue_date.strftime("%d %b %Y"), body_style),
        ],
        [
            Paragraph("<b>File reference</b>", body_style),
            Paragraph(html.escape(event_file.file_reference or "—"), body_style),
            Paragraph("<b>Valid until</b>", body_style),
            Paragraph(proposal.valid_until.strftime("%d %b %Y"), body_style),
        ],
        [
            Paragraph("<b>Event</b>", body_style),
            Paragraph(html.escape(event.exhibition_name), body_style),
            Paragraph("<b>Transport</b>", body_style),
            Paragraph(html.escape(event_file.transport_type), body_style),
        ],
        [
            Paragraph("<b>Status</b>", body_style),
            Paragraph(html.escape(proposal.status), body_style),
            Paragraph("<b>Master ref.</b>", body_style),
            Paragraph(html.escape(event.master_file_ref or "—"), body_style),
        ],
    ]
    story.append(
        Table(
            quote_details,
            colWidths=[27 * mm, 61 * mm, 27 * mm, 61 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F8FA")),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E3E8")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            ),
        )
    )
    story.extend(
        [
            Spacer(1, 6 * mm),
            Paragraph("<b>Prepared for</b>", styles["Heading3"]),
            Paragraph(html.escape(proposal.recipient_name), body_style),
        ]
    )
    if proposal.recipient_contact:
        story.append(
            Paragraph(f"Attn: {html.escape(proposal.recipient_contact)}", body_style)
        )
    if proposal.recipient_phone:
        story.append(Paragraph(f"Phone: {html.escape(proposal.recipient_phone)}", body_style))
    if proposal.recipient_email:
        story.append(Paragraph(html.escape(proposal.recipient_email), body_style))
    if proposal.recipient_address:
        escaped_address = html.escape(proposal.recipient_address).replace("\n", "<br/>")
        story.append(Paragraph(escaped_address, body_style))

    goods_by_direction = {
        direction: [item for item in event_file.goods_items if item.direction == direction]
        for direction in GOODS_DIRECTIONS
    }

    def cargo_summary(items):
        pieces = sum(item.pieces for item in items)
        weight = sum((Decimal(item.weight_kg or 0) for item in items), Decimal("0"))
        volume = sum(
            (
                Decimal(item.length_cm)
                * Decimal(item.width_cm)
                * Decimal(item.height_cm)
                * item.pieces
                / Decimal("1000000")
                for item in items
            ),
            Decimal("0"),
        )
        return f"{pieces} pcs · {weight:,.2f} kg · {volume:,.3f} CBM"

    cargo_rows = [
        [
            Paragraph("<b>Linked shipment file</b>", body_style),
            Paragraph(html.escape(event_file.file_reference or "—"), body_style),
            Paragraph("<b>Exhibitor</b>", body_style),
            Paragraph(
                html.escape(
                    event_file.exhibitor_contact.company_name
                    if event_file.exhibitor_contact
                    else "—"
                ),
                body_style,
            ),
        ],
        [
            Paragraph("<b>Venue</b>", body_style),
            Paragraph(html.escape(event.venue or "—"), body_style),
            Paragraph("<b>Hall / booth</b>", body_style),
            Paragraph(
                html.escape(
                    " / ".join(
                        part
                        for part in (event_file.hall_number, event_file.booth_number)
                        if part
                    )
                    or "—"
                ),
                body_style,
            ),
        ],
    ]
    route = ""
    if event_file.air_details:
        route_parts = [
            event_file.air_details.hawb_departure,
            event_file.air_details.hawb_arrival,
        ]
        if any(route_parts):
            route = " to ".join(part or "—" for part in route_parts)
    if route:
        cargo_rows.append(
            [
                Paragraph("<b>Air route</b>", body_style),
                Paragraph(html.escape(route), body_style),
                Paragraph("<b>Shipment</b>", body_style),
                Paragraph("Air", body_style),
            ]
        )
    for direction, title in (("export", "Export cargo"), ("import", "Import cargo")):
        if goods_by_direction[direction]:
            cargo_rows.append(
                [
                    Paragraph(f"<b>{title}</b>", body_style),
                    Paragraph(html.escape(cargo_summary(goods_by_direction[direction])), body_style),
                    Paragraph("", body_style),
                    Paragraph("", body_style),
                ]
            )
    story.extend(
        [
            Spacer(1, 5 * mm),
            Paragraph("<b>Event and shipment details</b>", styles["Heading3"]),
            Table(
                cargo_rows,
                colWidths=[29 * mm, 59 * mm, 27 * mm, 61 * mm],
                style=TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F3F8FA")),
                        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#F3F8FA")),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E3E8")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 7),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ]
                ),
            ),
            Spacer(1, 6 * mm),
        ]
    )

    currency_label = html.escape(proposal.currency)
    line_rows = [[
        Paragraph("<b>Description</b>", body_style),
        Paragraph("<b>Qty</b>", body_style),
        Paragraph("<b>Unit</b>", body_style),
        Paragraph(f"<b>Rate ({currency_label})</b>", body_style),
        Paragraph(f"<b>Amount ({currency_label})</b>", body_style),
    ]]
    for item in proposal.items:
        line_rows.append(
            [
                Paragraph(html.escape(item.description), body_style),
                f"{item.quantity:,.3f}".rstrip("0").rstrip("."),
                html.escape(item.unit or "—"),
                f"{item.unit_price:,.2f}",
                f"{item.line_total:,.2f}",
            ]
        )
    items_table = Table(
        line_rows,
        colWidths=[78 * mm, 18 * mm, 20 * mm, 29 * mm, 31 * mm],
        repeatRows=1,
        hAlign="LEFT",
    )
    items_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183541")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E3E8")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F6FAFB")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(items_table)

    totals = [
        ["Subtotal", f"{currency_label} {proposal.subtotal:,.2f}"],
    ]
    if proposal.discount_amount:
        discount_label = (
            f"Agency Discount ({proposal.discount_percent:.2f}%)"
            if proposal.discount_type == "percent"
            else "Agency Discount"
        )
        totals.append(
            [
                discount_label,
                f"- {currency_label} {proposal.discount_amount:,.2f}",
            ]
        )
    totals.append(
        [
            f"{proposal.tax_name or 'Tax'} ({proposal.tax_percent:.2f}%)",
            f"{currency_label} {proposal.tax_amount:,.2f}",
        ]
    )
    totals.append(["Total", f"{currency_label} {proposal.total:,.2f}"])
    totals_table = Table(totals, colWidths=[135 * mm, 41 * mm], hAlign="RIGHT")
    totals_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAF6FA")),
                ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor("#183541")),
                ("LINEABOVE", (0, -1), (-1, -1), 1, colors.HexColor("#08a5d2")),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.extend([Spacer(1, 4 * mm), totals_table])
    if proposal.bank_name and proposal.bank_account_number:
        bank_rows = [
            [
                Paragraph("<b>Beneficiary</b>", body_style),
                Paragraph(html.escape(proposal.bank_beneficiary), body_style),
                Paragraph("<b>Bank</b>", body_style),
                Paragraph(html.escape(proposal.bank_name), body_style),
            ],
            [
                Paragraph("<b>Account no.</b>", body_style),
                Paragraph(html.escape(proposal.bank_account_number), body_style),
                Paragraph("<b>Account type / currency</b>", body_style),
                Paragraph(
                    html.escape(
                        " / ".join(
                            part
                            for part in (proposal.bank_account_type, proposal.bank_currency)
                            if part
                        )
                        or "—"
                    ),
                    body_style,
                ),
            ],
        ]
        if proposal.bank_branch or proposal.bank_ifsc or proposal.bank_swift:
            bank_rows.append(
                [
                    Paragraph("<b>Branch</b>", body_style),
                    Paragraph(html.escape(proposal.bank_branch or "—"), body_style),
                    Paragraph("<b>IFSC / SWIFT</b>", body_style),
                    Paragraph(
                        html.escape(
                            " / ".join(
                                part for part in (proposal.bank_ifsc, proposal.bank_swift) if part
                            )
                            or "—"
                        ),
                        body_style,
                    ),
                ]
            )
        story.extend(
            [
                Spacer(1, 7 * mm),
                Paragraph(
                    f"<b>Payment details · {html.escape(proposal.bank_label)}</b>",
                    styles["Heading3"],
                ),
                Table(
                    bank_rows,
                    colWidths=[29 * mm, 59 * mm, 27 * mm, 61 * mm],
                    style=TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F3F8FA")),
                            ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#F3F8FA")),
                            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E3E8")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 7),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                            ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                        ]
                    ),
                ),
            ]
        )
    if proposal.terms:
        story.extend(
            [
                Spacer(1, 7 * mm),
                Paragraph("<b>Terms &amp; conditions</b>", styles["Heading3"]),
                Paragraph(html.escape(proposal.terms).replace("\n", "<br/>"), body_style),
            ]
        )
    if proposal.notes:
        story.extend(
            [
                Spacer(1, 4 * mm),
                Paragraph("<b>Notes</b>", styles["Heading3"]),
                Paragraph(html.escape(proposal.notes).replace("\n", "<br/>"), body_style),
            ]
        )
    story.extend(
        [
            Spacer(1, 8 * mm),
            Paragraph(
                "Thank you for the opportunity to submit this quotation.",
                small_style,
            ),
        ]
    )
    document.build(story)
    output.seek(0)
    return output


def delivery_note_contact_details(contact):
    details = proposal_recipient_details(contact)
    profile = contact.agent_profile if contact and contact.kind == "agent" else None
    details["country"] = profile.country if profile else ""
    return details


def event_file_delivery_note_data(file_record):
    cargo_rows = []
    total_pieces = 0
    total_weight = Decimal("0")
    total_cbm_exact = Decimal("0")
    for item in file_record.goods_items:
        if item.direction != "import":
            continue
        metrics = calculate_event_goods_metrics(
            {
                "length_cm": item.length_cm,
                "width_cm": item.width_cm,
                "height_cm": item.height_cm,
                "pieces": item.pieces,
            },
            file_record.transport_type,
        )
        total_pieces += item.pieces
        total_weight += item.weight_kg
        total_cbm_exact += (
            item.length_cm * item.width_cm * item.height_cm * item.pieces
        ) / Decimal("1000000")
        cargo_rows.append(
            {
                "marking": item.marking,
                "pieces": item.pieces,
                "pack_type": item.pack_other if item.pack_type == "Other" else item.pack_type,
                "length_cm": item.length_cm,
                "width_cm": item.width_cm,
                "height_cm": item.height_cm,
                "weight_kg": item.weight_kg,
                "description": item.description,
                "cbm": metrics["cbm"],
            }
        )
    air_details = file_record.air_details
    return {
        "exhibitor": delivery_note_contact_details(file_record.exhibitor_contact),
        "customer": delivery_note_contact_details(file_record.customer_contact),
        "agent": delivery_note_contact_details(file_record.agent_contact),
        "cargo_rows": cargo_rows,
        "total_pieces": total_pieces,
        "total_weight_kg": total_weight.quantize(Decimal("0.01")),
        "total_cbm": total_cbm_exact.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
        "tracking_reference": " / ".join(
            value
            for value in (
                air_details.mawb_no if air_details else "",
                air_details.hawb_no if air_details else "",
            )
            if value
        ),
    }


def build_event_file_delivery_note_pdf(file_record, note_data, form_values, delivery_date):
    output = BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Delivery Note {file_record.file_reference}",
        author="R.E. Rogers India Pvt. Ltd.",
    )
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle(
        "DeliveryBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=10,
        splitLongWords=True,
    )
    small_style = ParagraphStyle(
        "DeliverySmall",
        parent=body_style,
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor("#5E6F78"),
    )
    heading_style = ParagraphStyle(
        "DeliveryHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=17,
        alignment=2,
        textColor=colors.HexColor("#183541"),
    )
    section_style = ParagraphStyle(
        "DeliverySection",
        parent=styles["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#183541"),
    )
    header_label_style = ParagraphStyle(
        "DeliveryHeaderLabel",
        parent=small_style,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#52636C"),
    )
    header_value_style = ParagraphStyle(
        "DeliveryHeaderValue",
        parent=body_style,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#183541"),
    )
    table_header_style = ParagraphStyle(
        "DeliveryTableHeader",
        parent=small_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )

    def safe_paragraph(value, style=body_style):
        return Paragraph(html.escape(str(value or "—")).replace("\n", "<br/>"), style)

    logo_path = BASE_DIR / "static" / "rogers-40-years-logo.jpeg"
    logo_cell = (
        PDFImage(str(logo_path), width=48 * mm, height=18 * mm, kind="proportional")
        if logo_path.is_file()
        else Paragraph("<b>R.E. ROGERS INDIA PVT. LTD.</b>", section_style)
    )
    company_details = Paragraph(
        "<b>R.E. Rogers India Pvt. Ltd. · Delhi Office</b><br/>"
        f"{html.escape(PROPOSAL_COMPANY_ADDRESS)}<br/>"
        f"GSTIN: {html.escape(PROPOSAL_COMPANY_GSTIN)}<br/>"
        f"{html.escape(PROPOSAL_COMPANY_EMAIL)} · "
        f"{html.escape(PROPOSAL_COMPANY_WEBSITE)}",
        small_style,
    )
    event = file_record.event
    story = [
        Table(
            [[logo_cell, Paragraph("DELIVERY NOTE / POD", heading_style)]],
            colWidths=[92 * mm, 94 * mm],
            style=TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                    ("LINEBELOW", (0, 0), (-1, -1), 1.2, colors.HexColor("#08A5D2")),
                ]
            ),
        ),
        Spacer(1, 3 * mm),
        company_details,
        Spacer(1, 5 * mm),
    ]

    info_rows = [
        [
            safe_paragraph("EXHIBITION", header_label_style),
            safe_paragraph("VENUE / LOCATION", header_label_style),
            safe_paragraph("HALL", header_label_style),
            safe_paragraph("BOOTH", header_label_style),
        ],
        [
            safe_paragraph(event.exhibition_name, header_value_style),
            safe_paragraph(event.venue, header_value_style),
            safe_paragraph(file_record.hall_number, header_value_style),
            safe_paragraph(file_record.booth_number, header_value_style),
        ],
        [
            safe_paragraph("DELIVERY DATE", header_label_style),
            safe_paragraph("FILE REFERENCE", header_label_style),
            safe_paragraph("TRANSPORT MODE", header_label_style),
            safe_paragraph("AWB / TRACKING REF", header_label_style),
        ],
        [
            safe_paragraph(delivery_date.strftime("%d %b %Y"), header_value_style),
            safe_paragraph(file_record.file_reference, header_value_style),
            safe_paragraph(
                file_record.custom_transport
                if file_record.transport_type == "Other"
                else file_record.transport_type,
                header_value_style,
            ),
            safe_paragraph(note_data["tracking_reference"], header_value_style),
        ],
    ]
    story.extend(
        [
            Table(
                info_rows,
                colWidths=[46.5 * mm] * 4,
                style=TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F7F9")),
                        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D7E1E5")),
                        ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D7E1E5")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 6),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ]
                ),
            ),
            Spacer(1, 5 * mm),
        ]
    )

    def party_paragraphs(label, party, include_country=False):
        rows = [Paragraph(f"<b>{html.escape(label.upper())}</b>", section_style)]
        if party["company_name"]:
            rows.append(safe_paragraph(party["company_name"], header_value_style))
        if party["contact_person_name"]:
            rows.append(safe_paragraph(f"Contact: {party['contact_person_name']}", small_style))
        if party["address"]:
            rows.append(safe_paragraph(party["address"], small_style))
        if include_country and party["country"]:
            rows.append(safe_paragraph(party["country"], small_style))
        contact_line = " · ".join(value for value in (party["phone"], party["email"]) if value)
        if contact_line:
            rows.append(safe_paragraph(contact_line, small_style))
        return rows

    exhibitor_party = note_data["exhibitor"]
    if not exhibitor_party["company_name"]:
        exhibitor_party = note_data["customer"]
    party_width = 62 * mm
    story.extend(
        [
            Table(
                [[
                    party_paragraphs("Exhibitor / Customer", exhibitor_party),
                    party_paragraphs("Agent", note_data["agent"], include_country=True),
                    [
                        Paragraph("<b>SHIPMENT TRACKING</b>", section_style),
                        safe_paragraph(
                            f"MAWB / HAWB: {note_data['tracking_reference'] or '—'}"
                        ),
                        safe_paragraph(
                            f"File date: {file_record.file_date.strftime('%d %b %Y')}"
                            if file_record.file_date
                            else "File date: —"
                        ),
                    ],
                ]],
                colWidths=[party_width, party_width, party_width],
                style=TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5FAFC")),
                        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D7E7ED")),
                        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D7E7ED")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 7),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                        ("TOPPADDING", (0, 0), (-1, -1), 7),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                    ]
                ),
            ),
            Spacer(1, 5 * mm),
            Paragraph("<b>CARGO SPECIFICATIONS &amp; PACKAGE DETAILS</b>", section_style),
            Spacer(1, 2 * mm),
        ]
    )

    cargo_table_rows = [[
        safe_paragraph("Marks & Nos", table_header_style),
        safe_paragraph("No. of Pcs", table_header_style),
        safe_paragraph("Packing", table_header_style),
        safe_paragraph("Dimensions (L × W × H cm)", table_header_style),
        safe_paragraph("Gross Weight", table_header_style),
        safe_paragraph("CBM", table_header_style),
        safe_paragraph("Goods Description", table_header_style),
    ]]
    for item in note_data["cargo_rows"]:
        cargo_table_rows.append(
            [
                safe_paragraph(item["marking"]),
                safe_paragraph(item["pieces"]),
                safe_paragraph(item["pack_type"]),
                safe_paragraph(
                    f"{item['length_cm']} × {item['width_cm']} × {item['height_cm']}"
                ),
                safe_paragraph(f"{item['weight_kg']} kg"),
                safe_paragraph(item["cbm"]),
                safe_paragraph(item["description"]),
            ]
        )
    cargo_table_rows.append(
        [
            Paragraph("<b>TOTAL</b>", body_style),
            Paragraph(f"<b>{note_data['total_pieces']}</b>", body_style),
            "",
            "",
            Paragraph(f"<b>{note_data['total_weight_kg']} kg</b>", body_style),
            Paragraph(f"<b>{note_data['total_cbm']} CBM</b>", body_style),
            "",
        ]
    )
    story.extend(
        [
            Table(
                cargo_table_rows,
                colWidths=[28 * mm, 14 * mm, 20 * mm, 39 * mm, 23 * mm, 16 * mm, 46 * mm],
                repeatRows=1,
                style=TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183541")),
                        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAF3F6")),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CED9DE")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("ALIGN", (1, 1), (1, -1), "CENTER"),
                        ("ALIGN", (4, 1), (5, -1), "RIGHT"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ]
                ),
            ),
            Spacer(1, 4 * mm),
            Paragraph("<b>CONDITION AT DELIVERY</b>", section_style),
            safe_paragraph(form_values["condition_at_delivery"]),
            Spacer(1, 5 * mm),
        ]
    )

    receiver_name = form_values["receiver_name"] or " "
    receiver_fields = [
        ("Receiver Full Name", receiver_name),
        ("Designation", form_values["receiver_designation"]),
        ("Contact No.", form_values["receiver_phone"]),
        ("Email Address", form_values["receiver_email"]),
    ]
    receiver_content = [Paragraph("<b>RECEIVED BY (EXHIBITOR REPRESENTATIVE)</b>", section_style)]
    for label, value in receiver_fields:
        receiver_content.extend(
            [
                safe_paragraph(label, small_style),
                safe_paragraph(value),
                Spacer(1, 2 * mm),
            ]
        )
    receiver_content.extend(
        [
            Spacer(1, 5 * mm),
            Paragraph("____________________________________", small_style),
            safe_paragraph("Receiver signature, name & company stamp", small_style),
        ]
    )

    handover_content = [
        Paragraph("<b>DELIVERED / HANDED OVER BY</b>", section_style),
        safe_paragraph("Company: R.E. Rogers India Pvt. Ltd."),
        safe_paragraph(f"R.E. Rogers representative: {form_values['supervisor_name']}"),
        safe_paragraph(f"Designation: {form_values['supervisor_designation']}"),
        safe_paragraph(f"Contact: {form_values['supervisor_phone']}"),
        Spacer(1, 15 * mm),
        Paragraph("____________________________________", small_style),
        safe_paragraph("Authorized signature (Rogers supervisor)", small_style),
    ]
    story.append(
        Table(
            [[receiver_content, handover_content]],
            colWidths=[93 * mm, 93 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#F2F8FB")),
                    ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#F7F9FA")),
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D7E1E5")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7E1E5")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 9),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                    ("TOPPADDING", (0, 0), (-1, -1), 9),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                ]
            ),
        )
    )
    document.build(story)
    output.seek(0)
    return output


def proposal_recipient_details(contact):
    if contact is None:
        return {
            "kind": "",
            "company_name": "",
            "contact_person_name": "",
            "address": "",
            "email": "",
            "phone": "",
        }

    profile = contact.agent_profile if contact.kind == "agent" else None
    address = contact.address
    contact_person = contact.contact_person_name
    email = contact.email
    phone = contact.phone
    if profile:
        address = "\n".join(
            part
            for part in (
                profile.address,
                ", ".join(part for part in (profile.city, profile.postal_code) if part),
                profile.country,
            )
            if part
        )
        if not contact_person and profile.contact_people:
            contact_person = profile.contact_people[0].contact_name
        email = profile.company_email or email
        phone = profile.company_phone or phone
    return {
        "kind": contact.kind,
        "company_name": contact.company_name,
        "contact_person_name": contact_person,
        "address": address,
        "email": email,
        "phone": phone,
    }


def event_file_proposal_template_data(file_record):
    directory = Contact.query.filter(Contact.kind.in_(("customer", "agent"))).order_by(
        Contact.company_name.asc()
    )
    return {
        "recipient_options": [
            proposal_recipient_details(contact) for contact in directory.all()
        ],
        "file_recipient_options": {
            "customer": proposal_recipient_details(file_record.customer_contact),
            "agent": proposal_recipient_details(file_record.agent_contact),
        },
    }


def event_file_proposal_form_defaults(file_record, proposal=None):
    if proposal is not None:
        return {
            "recipient_type": proposal.recipient_type,
            "recipient_name": proposal.recipient_name,
            "recipient_contact": proposal.recipient_contact,
            "recipient_phone": proposal.recipient_phone,
            "recipient_email": proposal.recipient_email,
            "recipient_address": proposal.recipient_address,
            "issue_date": proposal.issue_date.isoformat(),
            "valid_until": proposal.valid_until.isoformat(),
            "status": proposal.status,
            "currency": proposal.currency,
            "discount_type": proposal.discount_type,
            "discount_value": str(proposal.discount_value),
            "tax_name": proposal.tax_name,
            "tax_percent": str(proposal.tax_percent),
            "terms": proposal.terms,
            "notes": proposal.notes,
            "bank_account_slot": str(proposal.bank_account_slot or ""),
        }

    recipient = file_record.customer_contact or file_record.agent_contact
    recipient_type = "customer" if file_record.customer_contact else "agent"
    recipient_details = proposal_recipient_details(recipient)
    today = date.today()
    return {
        "recipient_type": recipient_type,
        "recipient_name": recipient_details["company_name"],
        "recipient_contact": recipient_details["contact_person_name"],
        "recipient_phone": recipient_details["phone"],
        "recipient_email": recipient_details["email"],
        "recipient_address": recipient_details["address"],
        "issue_date": today.isoformat(),
        "valid_until": (today + timedelta(days=30)).isoformat(),
        "status": "DRAFT",
        "currency": "INR",
        "discount_type": "amount",
        "discount_value": "0",
        "tax_name": "",
        "tax_percent": "",
        "terms": "",
        "notes": "",
        "bank_account_slot": "",
    }


def event_file_proposal_item_rows(proposal=None):
    if proposal is None or not proposal.items:
        return [{"description": "", "quantity": "", "unit": "", "unit_price": ""}]
    return [
        {
            "description": item.description,
            "quantity": str(item.quantity),
            "unit": item.unit,
            "unit_price": str(item.unit_price),
        }
        for item in proposal.items
    ]


def event_file_proposal_bank_slots():
    stored_accounts = {
        account.slot: account
        for account in ProposalBankAccount.query.order_by(ProposalBankAccount.slot.asc()).all()
    }
    return [
        SimpleNamespace(
            slot=slot,
            label=account.label if account else f"Bank option {slot}",
            beneficiary=account.beneficiary if account else "",
            bank_name=account.bank_name if account else "",
            account_number=account.account_number if account else "",
            account_type=account.account_type if account else "",
            currency=account.currency if account else "",
            branch=account.branch if account else "",
            ifsc=account.ifsc if account else "",
            swift=account.swift if account else "",
            is_configured=account.is_configured if account else False,
        )
        for slot in (1, 2)
        for account in (stored_accounts.get(slot),)
    ]


def event_file_proposal_request_state(form):
    form_values = {
        field: form.get(field, "").strip()
        for field in (
            "recipient_type",
            "recipient_name",
            "recipient_contact",
            "recipient_phone",
            "recipient_email",
            "recipient_address",
            "issue_date",
            "valid_until",
            "status",
            "currency",
            "discount_type",
            "discount_value",
            "tax_name",
            "tax_percent",
            "terms",
            "notes",
            "bank_account_slot",
        )
    }
    if not form_values["discount_value"] and form.get("discount_percent"):
        form_values["discount_value"] = form.get("discount_percent", "").strip()
        form_values["discount_type"] = "percent"

    descriptions = form.getlist("item_description")
    quantities = form.getlist("item_quantity")
    units = form.getlist("item_unit")
    unit_prices = form.getlist("item_unit_price")
    row_count = max(len(descriptions), len(quantities), len(units), len(unit_prices))
    item_rows = [
        {
            "description": descriptions[index] if index < len(descriptions) else "",
            "quantity": quantities[index] if index < len(quantities) else "",
            "unit": units[index] if index < len(units) else "",
            "unit_price": unit_prices[index] if index < len(unit_prices) else "",
        }
        for index in range(min(row_count, 20))
    ]
    if not item_rows:
        item_rows = [{"description": "", "quantity": "", "unit": "", "unit_price": ""}]

    return form_values, item_rows, row_count


def parse_event_file_proposal_request(
    form,
    files,
    bank_accounts,
    existing_proposal=None,
    form_state=None,
):
    form_values, item_rows, row_count = (
        form_state if form_state is not None else event_file_proposal_request_state(form)
    )
    if row_count > 20:
        raise ValueError("A quotation can contain up to 20 line items.")
    if form_values["recipient_type"] not in {"customer", "agent"}:
        raise ValueError("Choose whether this quotation is for the customer or agent.")
    recipient_name = form_values["recipient_name"][:200]
    if not recipient_name:
        raise ValueError("Enter the customer or agent company name.")
    recipient_contact = form_values["recipient_contact"][:200]
    recipient_phone = form_values["recipient_phone"][:50]
    recipient_email = form_values["recipient_email"][:254]
    if recipient_email and (
        "@" not in recipient_email or "." not in recipient_email.rsplit("@", 1)[-1]
    ):
        raise ValueError("Enter a valid recipient email address.")
    recipient_address = form_values["recipient_address"][:2000]
    supported_currencies = {code for code, _ in PROPOSAL_CURRENCY_OPTIONS}
    currency = form_values["currency"].upper()
    if currency not in supported_currencies:
        raise ValueError("Choose a supported quotation currency.")
    status = form_values["status"].upper()
    if status not in PROPOSAL_STATUS_OPTIONS:
        raise ValueError("Choose a valid quotation status.")
    try:
        issue_date = date.fromisoformat(form_values["issue_date"])
        valid_until = date.fromisoformat(form_values["valid_until"])
    except ValueError as error:
        raise ValueError("Enter valid quotation issue and expiry dates.") from error
    if valid_until < issue_date:
        raise ValueError("The quotation expiry date cannot be before its issue date.")

    discount_type = form_values["discount_type"]
    if discount_type not in {"amount", "percent"}:
        raise ValueError("Choose a valid discount type.")
    discount_value = parse_proposal_decimal(
        form_values["discount_value"],
        "Agency discount",
        default=Decimal("0"),
        minimum=Decimal("0"),
        maximum=(
            Decimal("100")
            if discount_type == "percent"
            else Decimal("99999999999999.99")
        ),
    )
    tax_percent = parse_proposal_decimal(
        form_values["tax_percent"],
        "Tax percentage",
        default=Decimal("0"),
        minimum=Decimal("0"),
        maximum=Decimal("100"),
    )
    tax_name = form_values["tax_name"][:60]
    terms = form_values["terms"][:4000]
    notes = form_values["notes"][:2000]

    parsed_items = []
    subtotal = Decimal("0.00")
    for index, row in enumerate(item_rows):
        description = row["description"].strip()
        quantity_value = row["quantity"].strip()
        unit = row["unit"].strip()[:40]
        rate_value = row["unit_price"].strip()
        if not any((description, quantity_value, unit, rate_value)):
            continue
        if not description:
            raise ValueError(f"Line item {index + 1} needs a description.")
        if len(description) > 250:
            raise ValueError(f"Line item {index + 1} description is too long.")
        quantity = parse_proposal_decimal(
            quantity_value,
            f"Line item {index + 1} quantity",
            minimum=Decimal("0.001"),
            maximum=Decimal("999999999.999"),
        )
        unit_price = parse_proposal_decimal(
            rate_value,
            f"Line item {index + 1} rate",
            minimum=Decimal("0"),
            maximum=Decimal("999999999999.99"),
        )
        if unit_price != unit_price.quantize(Decimal("0.01")):
            raise ValueError(f"Line item {index + 1} rate can have up to 2 decimal places.")
        line_total = (quantity * unit_price).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
        if line_total > Decimal("99999999999999.99"):
            raise ValueError(f"Line item {index + 1} amount is too large.")
        subtotal += line_total
        parsed_items.append(
            {
                "description": description,
                "quantity": quantity,
                "unit": unit,
                "unit_price": unit_price,
                "line_total": line_total,
            }
        )
    if not parsed_items:
        raise ValueError("Add at least one quotation line item.")
    if subtotal > Decimal("99999999999999.99"):
        raise ValueError("The quotation subtotal is too large.")

    if discount_type == "percent":
        discount_percent = discount_value
        discount_amount = (subtotal * discount_percent / Decimal("100")).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    else:
        discount_percent = Decimal("0")
        discount_amount = discount_value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if discount_amount > subtotal:
        raise ValueError("The agency discount cannot exceed the quotation subtotal.")
    taxable_amount = subtotal - discount_amount
    tax_amount = (taxable_amount * tax_percent / Decimal("100")).quantize(
        Decimal("0.01"),
        rounding=ROUND_HALF_UP,
    )
    total = taxable_amount + tax_amount
    if total > Decimal("99999999999999.99"):
        raise ValueError("The quotation total is too large.")

    bank_slot = None
    bank_snapshot = {
        "bank_label": "",
        "bank_beneficiary": "",
        "bank_name": "",
        "bank_account_number": "",
        "bank_account_type": "",
        "bank_currency": "",
        "bank_branch": "",
        "bank_ifsc": "",
        "bank_swift": "",
    }
    requested_bank = form_values["bank_account_slot"]
    if requested_bank == "saved" and existing_proposal and existing_proposal.bank_name:
        bank_slot = existing_proposal.bank_account_slot
        bank_snapshot = {
            field: getattr(existing_proposal, field)
            for field in bank_snapshot
        }
    elif requested_bank:
        try:
            bank_slot = int(requested_bank)
        except ValueError as error:
            raise ValueError("Choose a valid bank account option.") from error
        account = next((entry for entry in bank_accounts if entry.slot == bank_slot), None)
        if account is None or not account.is_configured:
            raise ValueError("Configure the selected bank option before using it.")
        bank_snapshot = {
            "bank_label": account.label,
            "bank_beneficiary": account.beneficiary,
            "bank_name": account.bank_name,
            "bank_account_number": account.account_number,
            "bank_account_type": account.account_type,
            "bank_currency": account.currency,
            "bank_branch": account.branch,
            "bank_ifsc": account.ifsc,
            "bank_swift": account.swift,
        }

    upload = files.get("quotation_upload")
    if upload is not None and upload.filename:
        uploaded_filename, uploaded_mimetype, uploaded_data = read_proposal_upload(upload)
    elif existing_proposal and existing_proposal.uploaded_data and not form.get("remove_uploaded"):
        uploaded_filename = existing_proposal.uploaded_filename
        uploaded_mimetype = existing_proposal.uploaded_mimetype
        uploaded_data = existing_proposal.uploaded_data
    else:
        uploaded_filename, uploaded_mimetype, uploaded_data = "", "", None

    return {
        "recipient_type": form_values["recipient_type"],
        "recipient_name": recipient_name,
        "recipient_contact": recipient_contact,
        "recipient_phone": recipient_phone,
        "recipient_email": recipient_email,
        "recipient_address": recipient_address,
        "currency": currency,
        "issue_date": issue_date,
        "valid_until": valid_until,
        "status": status,
        "discount_type": discount_type,
        "discount_value": discount_value,
        "discount_percent": discount_percent,
        "tax_name": tax_name,
        "tax_percent": tax_percent,
        "subtotal": subtotal,
        "discount_amount": discount_amount,
        "tax_amount": tax_amount,
        "total": total,
        "terms": terms,
        "notes": notes,
        "bank_account_slot": bank_slot,
        **bank_snapshot,
        "uploaded_filename": uploaded_filename,
        "uploaded_mimetype": uploaded_mimetype,
        "uploaded_data": uploaded_data,
        "items": parsed_items,
        "form_values": form_values,
        "item_rows": item_rows,
    }


def apply_event_file_proposal_values(proposal, values):
    fields = (
        "recipient_type",
        "recipient_name",
        "recipient_contact",
        "recipient_phone",
        "recipient_email",
        "recipient_address",
        "currency",
        "issue_date",
        "valid_until",
        "status",
        "discount_type",
        "discount_value",
        "discount_percent",
        "tax_name",
        "tax_percent",
        "subtotal",
        "discount_amount",
        "tax_amount",
        "total",
        "terms",
        "notes",
        "uploaded_filename",
        "uploaded_mimetype",
        "uploaded_data",
        "bank_account_slot",
        "bank_label",
        "bank_beneficiary",
        "bank_name",
        "bank_account_number",
        "bank_account_type",
        "bank_currency",
        "bank_branch",
        "bank_ifsc",
        "bank_swift",
    )
    for field in fields:
        setattr(proposal, field, values[field])
    proposal.items.clear()
    proposal.items.extend(
        EventFileProposalItem(
            sort_order=index,
            description=item["description"],
            quantity=item["quantity"],
            unit=item["unit"],
            unit_price=item["unit_price"],
            line_total=item["line_total"],
        )
        for index, item in enumerate(values["items"])
    )
    proposal.updated_at = utc_now()


def create_pdf_report(records, report_title, filter_description, fields, heading_for_record, empty_message):
    output = BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title=report_title,
    )
    styles = getSampleStyleSheet()
    label_style = ParagraphStyle(
        "ShipmentLabel",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
    )
    value_style = ParagraphStyle(
        "ShipmentValue",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        splitLongWords=True,
    )

    generated_at = utc_now().astimezone(DISPLAY_TIMEZONE).strftime("%d-%m-%Y %I:%M:%S %p")
    story = [
        Paragraph(html.escape(report_title), styles["Title"]),
        Paragraph(
            f"Generated: {html.escape(generated_at)} IST | "
            f"Filter: {html.escape(filter_description)}",
            styles["BodyText"],
        ),
        Spacer(1, 8 * mm),
    ]

    if not records:
        story.append(Paragraph(html.escape(empty_message), styles["BodyText"]))
    for index, record in enumerate(records):
        if index:
            story.append(PageBreak())
        story.append(
            Paragraph(
                html.escape(heading_for_record(record)),
                styles["Heading2"],
            )
        )
        rows = []
        for field, label in fields:
            value = export_value(getattr(record, field))
            safe_value = html.escape(value).replace("\r\n", "\n").replace("\n", "<br/>")
            rows.append(
                [
                    Paragraph(html.escape(label), label_style),
                    Paragraph(safe_value or "—", value_style),
                ]
            )
        details = Table(rows, colWidths=[43 * mm, 143 * mm], splitByRow=1)
        details.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F0F4F8")),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9E2EC")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(details)

    document.build(story)
    output.seek(0)
    return output


def create_pdf_export(shipments, search, status_filter):
    filters = []
    if search:
        filters.append(f"Search: {search}")
    if status_filter:
        filters.append(f"Status: {status_filter}")
    return create_pdf_report(
        shipments,
        "Shipment Tracker Report",
        " | ".join(filters) or "All shipments",
        EXPORT_FIELDS,
        lambda shipment: f"{shipment.reference} — {shipment.tracking_status}",
        "No shipments match the selected filters.",
    )


def create_event_pdf_export(events, search):
    filter_description = f"Search: {search}" if search else "All events"
    return create_pdf_report(
        events,
        "Import Export Events Report",
        filter_description,
        EVENT_EXPORT_FIELDS,
        lambda event: event.exhibition_name,
        "No events match the selected search.",
    )


def create_event_file_pdf_export(files, search, selected_event):
    filters = []
    if search:
        filters.append(f"Search: {search}")
    if selected_event:
        filters.append(f"Event: {selected_event.exhibition_name}")
    return create_pdf_report(
        files,
        "Import Export Shipment Files Report",
        " | ".join(filters) or "All shipment files",
        EVENT_FILE_EXPORT_FIELDS,
        lambda file_record: f"{file_record.event.exhibition_name} - File #{file_record.id}",
        "No shipment files match the selected filters.",
    )


def format_duration(seconds):
    total_seconds = max(0, int(seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours}h {minutes}m {seconds}s"


def hash_session_token(token):
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def start_activity_session(user):
    token = secrets.token_urlsafe(32)
    now = utc_now()
    session["activity_session_token"] = token
    session.permanent = True
    db.session.add(
        UserSession(
            token_hash=hash_session_token(token),
            user_id=user.id,
            username=user.username,
            started_at=now,
            last_seen_at=now,
        )
    )
    db.session.commit()


def session_duration(session_record, now):
    start = as_utc(session_record.started_at)
    if session_record.ended_at:
        end = as_utc(session_record.ended_at)
    else:
        last_seen = as_utc(session_record.last_seen_at)
        end = min(now, last_seen + IDLE_TIMEOUT)
    return max(0, int((end - start).total_seconds()))


def close_idle_sessions(now=None):
    now = now or utc_now()
    cutoff = now - IDLE_TIMEOUT
    stale_sessions = UserSession.query.filter(
        UserSession.ended_at.is_(None),
        UserSession.last_seen_at < cutoff,
    ).all()
    for activity_session in stale_sessions:
        activity_session.ended_at = as_utc(activity_session.last_seen_at) + IDLE_TIMEOUT
        activity_session.end_reason = "idle_timeout"
    if stale_sessions:
        db.session.commit()


def import_legacy_shipments():
    if db.session.get(AppSetting, "legacy_shipments_imported"):
        return

    source_path = BASE_DIR / "shipments.json"
    if source_path.exists():
        try:
            with source_path.open("r", encoding="utf-8") as source_file:
                data = json.load(source_file)
        except (OSError, json.JSONDecodeError) as error:
            raise RuntimeError(f"Could not read legacy shipment data from {source_path}") from error
        if not isinstance(data, list):
            raise RuntimeError(f"Expected a list in {source_path}")

        seen_references = set()
        for item in data:
            if not isinstance(item, dict):
                raise RuntimeError(f"Invalid shipment record in {source_path}")
            reference = str(item.get("reference") or "").strip()
            reference_key = reference.casefold()
            if not reference:
                raise RuntimeError(f"A shipment in {source_path} has no reference")
            if reference_key in seen_references or Shipment.query.filter_by(reference_key=reference_key).first():
                raise RuntimeError(f"Duplicate shipment reference in {source_path}: {reference}")
            seen_references.add(reference_key)

            pending_items = item.get("pending_items") or []
            if isinstance(pending_items, list):
                pending_items = "\n".join(str(value) for value in pending_items)
            else:
                pending_items = str(pending_items)

            shipment = Shipment(
                reference=reference,
                reference_key=reference_key,
                agent_name=str(item.get("agent_name") or ""),
                client=str(item.get("client") or item.get("exhibitor_name") or "Unknown client"),
                exhibitor_name=str(item.get("exhibitor_name") or ""),
                exhibition_name=str(item.get("exhibition_name") or ""),
                supplier=str(item.get("supplier") or ""),
                origin=str(item.get("origin") or ""),
                destination=str(item.get("destination") or ""),
                tracking_status=str(item.get("tracking_status") or item.get("arrival_status") or "In Transit"),
                arrival_date=parse_legacy_date(item.get("arrival_date") or item.get("eta"), "arrival date"),
                arrival_time=None,
                expected_reexport_date=parse_legacy_date(
                    item.get("expected_reexport_date"),
                    "expected re-export date",
                ),
                shipment_value=str(item.get("shipment_value") or ""),
                stage=str(item.get("stage") or item.get("arrival_status") or "Received"),
                customs_status=str(item.get("customs_status") or "Not Filed"),
                customs_reference=str(item.get("customs_reference") or ""),
                payment_status=str(item.get("payment_status") or "Pending"),
                pending_items=pending_items,
                re_export_status=str(item.get("re_export_status") or "Not Planned"),
                re_export_reference=str(item.get("re_export_reference") or ""),
                container_no=str(item.get("container_no") or ""),
                warehouse_location=str(item.get("warehouse_location") or ""),
                notes=str(item.get("notes") or ""),
            )
            if shipment.arrival_date is None:
                raise RuntimeError(f"Shipment {reference} has no arrival date in {source_path}")
            db.session.add(shipment)

    db.session.add(AppSetting(key="legacy_shipments_imported", value="yes"))
    db.session.commit()


def make_shipment_form(shipment=None):
    form = ShipmentForm(obj=shipment)
    form.tracking_status.choices = [(status, status) for status in TRACKING_STATUSES]
    stage_choices = list(SHIPMENT_STAGES)
    if shipment and shipment.stage not in stage_choices:
        stage_choices.insert(0, shipment.stage)
    form.stage.choices = [(stage, stage) for stage in stage_choices]
    return form


def admin_required(view):
    @wraps(view)
    @login_required
    def wrapped_view(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return view(*args, **kwargs)

    return wrapped_view


def permission_required(permission):
    def decorator(view):
        @wraps(view)
        @login_required
        def wrapped_view(*args, **kwargs):
            if not current_user.has_permission(permission):
                abort(403)
            return view(*args, **kwargs)

        return wrapped_view

    return decorator


def is_safe_redirect_target(target):
    if not target:
        return False
    ref_url = urlsplit(request.host_url)
    test_url = urlsplit(urljoin(request.host_url, target))
    return test_url.scheme in {"http", "https"} and ref_url.netloc == test_url.netloc


def create_app(test_config=None):
    app = Flask(__name__, instance_path=str(BASE_DIR / "instance"))
    Path(app.instance_path).mkdir(parents=True, exist_ok=True)

    production = os.environ.get("APP_ENV", "").casefold() == "production"
    secret_key = os.environ.get("SECRET_KEY")
    if production and not secret_key:
        raise RuntimeError("SECRET_KEY must be configured before production startup.")

    default_database = f"sqlite:///{(Path(app.instance_path) / 'shipments.db').as_posix()}"
    database_url = os.environ.get("DATABASE_URL", default_database)
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql+psycopg://", 1)

    app.config.update(
        SECRET_KEY=secret_key or secrets.token_hex(32),
        SQLALCHEMY_DATABASE_URI=database_url,
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        WTF_CSRF_ENABLED=True,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=production,
        PERMANENT_SESSION_LIFETIME=IDLE_TIMEOUT,
        IMPORT_LEGACY_SHIPMENTS=True,
    )
    if test_config:
        app.config.update(test_config)
    if app.config["SQLALCHEMY_DATABASE_URI"].startswith("sqlite:"):
        app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
            "connect_args": {"check_same_thread": False},
        }

    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)

    with app.app_context():
        db.create_all()
        upgrade_user_permissions()
        upgrade_event_file_air_details()
        upgrade_agent_profile_location_fields()
        upgrade_contact_address()
        upgrade_event_file_proposal_fields()
        upgrade_event_file_delivery_note_history()
        if app.config["IMPORT_LEGACY_SHIPMENTS"]:
            import_legacy_shipments()

    @app.template_filter("ist")
    def format_ist(value):
        if value is None:
            return "—"
        return as_utc(value).astimezone(DISPLAY_TIMEZONE).strftime("%d-%m-%Y %I:%M:%S %p")

    @app.template_filter("duration")
    def format_duration_filter(seconds):
        return format_duration(seconds)

    @app.context_processor
    def proposal_company_profile():
        return {
            "proposal_company_branch": PROPOSAL_COMPANY_BRANCH,
            "proposal_company_address": PROPOSAL_COMPANY_ADDRESS,
            "proposal_company_gstin": PROPOSAL_COMPANY_GSTIN,
            "proposal_company_email": PROPOSAL_COMPANY_EMAIL,
            "proposal_company_website": PROPOSAL_COMPANY_WEBSITE,
        }

    @app.before_request
    def track_user_activity():
        if not current_user.is_authenticated or request.endpoint == "static":
            return

        now = utc_now()
        close_idle_sessions(now)
        token = session.get("activity_session_token")
        activity_session = None
        if token:
            activity_session = UserSession.query.filter_by(
                token_hash=hash_session_token(token),
                user_id=current_user.id,
            ).first()

        if activity_session and activity_session.ended_at:
            logout_user()
            session.pop("activity_session_token", None)
            flash("You were signed out after 15 minutes of inactivity. Please sign in again.", "warning")
            return redirect(url_for("login"))

        if activity_session is None:
            start_activity_session(current_user)
        elif now - as_utc(activity_session.last_seen_at) >= HEARTBEAT_INTERVAL:
            activity_session.last_seen_at = now
            db.session.commit()

    @app.get("/")
    def home():
        if current_user.is_authenticated:
            return redirect(url_for("departments"))
        return render_template("home.html", form=LoginForm(), next_url="")

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            return redirect(url_for("departments"))
        form = LoginForm()
        if form.validate_on_submit():
            username = form.username.data.strip()
            user = User.query.filter_by(username=username).first()
            if user and user.check_password(form.password.data):
                login_user(user)
                start_activity_session(user)
                flash("Login successful.", "success")
                target = request.form.get("next") or request.args.get("next")
                return redirect(target if is_safe_redirect_target(target) else url_for("departments"))
            flash("Incorrect username or password.", "error")
        next_url = request.form.get("next") or request.args.get("next", "")
        return render_template("home.html", form=form, next_url=next_url)

    @app.post("/logout")
    @login_required
    def logout():
        token = session.get("activity_session_token")
        if token:
            activity_session = UserSession.query.filter_by(
                token_hash=hash_session_token(token),
                user_id=current_user.id,
                ended_at=None,
            ).first()
            if activity_session:
                activity_session.ended_at = utc_now()
                activity_session.end_reason = "logout"
                db.session.commit()
        logout_user()
        session.clear()
        flash("You have been signed out.", "success")
        return redirect(url_for("home"))

    @app.route("/change-password", methods=["GET", "POST"])
    @login_required
    def change_password():
        form = ChangePasswordForm()
        if form.validate_on_submit():
            if not current_user.check_password(form.current_password.data):
                form.current_password.errors.append("The current password is incorrect.")
            else:
                current_user.set_password(form.new_password.data)
                db.session.commit()
                flash("Your password has been changed.", "success")
                return redirect(url_for("departments"))
        return render_template("change_password.html", form=form)

    @app.get("/activity/heartbeat")
    @login_required
    def activity_heartbeat():
        return "", 204

    @app.get("/departments")
    @permission_required("view")
    def departments():
        return render_template("departments.html", departments=DEPARTMENTS)

    @app.get("/departments/<department>/dashboard")
    @permission_required("view")
    def department_dashboard(department):
        department_info = get_department(department)
        session["active_department"] = department
        event_query = Event.query.filter_by(department=department)
        today = datetime.now(DISPLAY_TIMEZONE).date()
        return render_template(
            "department_dashboard.html",
            department=department_info,
            event_count=event_query.count(),
            upcoming_event_count=event_query.filter(Event.end_date >= today).count(),
            file_count=(
                EventShipmentFile.query.join(Event)
                .filter(Event.department == department)
                .count()
                if department == "import-export"
                else None
            ),
            shipment_count=Shipment.query.count() if department == "import-export" else None,
        )

    @app.get("/departments/<department>/proposals")
    @permission_required("view")
    def department_proposals(department):
        department_info = get_department(department)
        session["active_department"] = department
        if department == "import-export":
            proposals = (
                EventFileProposal.query.join(
                    EventShipmentFile,
                    EventFileProposal.file_id == EventShipmentFile.id,
                )
                .join(Event, EventShipmentFile.event_id == Event.id)
                .filter(Event.department == department)
                .order_by(EventFileProposal.created_at.desc())
                .all()
            )
            return render_template(
                "department_proposals.html",
                department=department_info,
                proposals=proposals,
            )
        return render_template(
            "module_placeholder.html",
            department=department_info,
            module_name="Proposal",
        )

    @app.get("/departments/<department>/invoicing")
    @login_required
    def department_invoicing(department):
        department_info = get_department(department)
        session["active_department"] = department
        return render_template(
            "module_placeholder.html",
            department=department_info,
            module_name="Invoicing",
        )

    @app.get("/departments/<department>/events")
    @permission_required("view")
    def department_events(department):
        department_info = get_department(department)
        session["active_department"] = department
        search = request.args.get("q", "").strip()
        configured = department == "import-export"
        records = []
        if configured:
            records = filtered_event_query(department, search).all()
        return render_template(
            "events.html",
            department=department_info,
            events=records,
            search=search,
            configured=configured,
        )

    @app.get("/departments/<department>/events.xlsx")
    @permission_required("view")
    def export_events_excel(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        search = request.args.get("q", "").strip()
        output = create_excel_export(
            filtered_event_query(department, search).all(),
            fields=EVENT_EXPORT_FIELDS,
            sheet_title="Events",
        )
        filename = f"events_{utc_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    @app.get("/departments/<department>/events.pdf")
    @permission_required("view")
    def export_events_pdf(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        search = request.args.get("q", "").strip()
        records = filtered_event_query(department, search).all()
        output = create_event_pdf_export(records, search)
        filename = f"events_{utc_now().strftime('%Y%m%d_%H%M%S')}.pdf"
        return send_file(
            output,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    @app.route("/departments/<department>/events/new", methods=["GET", "POST"])
    @permission_required("add")
    def add_event(department):
        department_info = get_department(department)
        if department != "import-export":
            abort(404)
        session["active_department"] = department
        form = EventForm()
        if (
            form.validate_on_submit()
            and event_date_ranges_valid(form)
            and event_master_reference_valid(form)
        ):
            event = Event(
                department=department,
                created_by_user_id=current_user.id,
            )
            form.populate_obj(event)
            event.department = department
            db.session.add(event)
            db.session.flush()
            ensure_event_file_sequences(event.id)
            snapshot = event_snapshot(event)
            audit_event(
                "create",
                "event",
                event.id,
                event.exhibition_name,
                {"fields": {key: {"before": None, "after": value} for key, value in snapshot.items()}},
            )
            db.session.commit()
            flash("Event saved.", "success")
            return redirect(url_for("department_events", department=department))
        return render_template(
            "event_form.html",
            form=form,
            department=department_info,
            page_title="Add event",
        )

    @app.route("/departments/<department>/events/<int:event_id>/edit", methods=["GET", "POST"])
    @permission_required("edit")
    def edit_event(department, event_id):
        department_info = get_department(department)
        if department != "import-export":
            abort(404)
        event = db.get_or_404(Event, event_id)
        if event.department != department:
            abort(404)
        session["active_department"] = department
        form = EventForm(obj=event)
        if (
            form.validate_on_submit()
            and event_date_ranges_valid(form)
            and event_master_reference_valid(form, event)
        ):
            before = event_snapshot(event)
            form.populate_obj(event)
            ensure_event_file_sequences(event.id)
            after = event_snapshot(event)
            if before != after:
                audit_event(
                    "edit",
                    "event",
                    event.id,
                    event.exhibition_name,
                    {
                        "fields": {
                            key: {"before": before[key], "after": after[key]}
                            for key in before
                            if before[key] != after[key]
                        }
                    },
                )
                db.session.commit()
                flash("Event updated.", "success")
            else:
                flash("No changes were made.", "warning")
            return redirect(url_for("department_events", department=department))
        return render_template(
            "event_form.html",
            form=form,
            department=department_info,
            page_title="Edit event",
        )

    @app.post("/departments/<department>/events/<int:event_id>/delete")
    @permission_required("delete")
    def delete_event(department, event_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        event = db.get_or_404(Event, event_id)
        if event.department != department:
            abort(404)
        snapshot = event_snapshot(event)
        audit_event(
            "delete",
            "event",
            event.id,
            event.exhibition_name,
            {"fields": {key: {"before": value, "after": None} for key, value in snapshot.items()}},
        )
        db.session.delete(event)
        db.session.commit()
        flash("Event deleted.", "success")
        return redirect(url_for("department_events", department=department))

    @app.get("/departments/<department>/files")
    @permission_required("view")
    def department_files(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        session["active_department"] = department
        search = request.args.get("q", "").strip()
        event_id = optional_query_integer("event_id")
        search_field = request.args.get("field", "all").strip()
        if search_field not in EVENT_FILE_SEARCH_FIELDS:
            search_field = "all"
        selected_event = None
        if event_id is not None:
            selected_event = db.get_or_404(Event, event_id)
            if selected_event.department != department:
                abort(404)
        records = filtered_event_file_query(department, search, event_id, search_field).all()
        events = Event.query.filter_by(department=department).order_by(Event.start_date.asc()).all()
        return render_template(
            "event_files.html",
            department=get_department(department),
            files=records,
            events=events,
            selected_event=selected_event,
            search=search,
            search_field=search_field,
            search_fields=EVENT_FILE_SEARCH_FIELDS,
            tracking_statuses=TRACKING_STATUSES,
        )

    @app.get("/departments/<department>/files.xlsx")
    @permission_required("view")
    def export_event_files_excel(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        search = request.args.get("q", "").strip()
        event_id = optional_query_integer("event_id")
        search_field = request.args.get("field", "all").strip()
        if search_field not in EVENT_FILE_SEARCH_FIELDS:
            search_field = "all"
        selected_event = db.get_or_404(Event, event_id) if event_id is not None else None
        if selected_event and selected_event.department != department:
            abort(404)
        output = create_excel_export(
            filtered_event_file_query(department, search, event_id, search_field).all(),
            fields=EVENT_FILE_EXPORT_FIELDS,
            sheet_title="Shipment Files",
        )
        filename = f"shipment_files_{utc_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    @app.get("/departments/<department>/files.pdf")
    @permission_required("view")
    def export_event_files_pdf(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        search = request.args.get("q", "").strip()
        event_id = optional_query_integer("event_id")
        search_field = request.args.get("field", "all").strip()
        if search_field not in EVENT_FILE_SEARCH_FIELDS:
            search_field = "all"
        selected_event = db.get_or_404(Event, event_id) if event_id is not None else None
        if selected_event and selected_event.department != department:
            abort(404)
        records = filtered_event_file_query(department, search, event_id, search_field).all()
        output = create_event_file_pdf_export(records, search, selected_event)
        filename = f"shipment_files_{utc_now().strftime('%Y%m%d_%H%M%S')}.pdf"
        return send_file(
            output,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    @app.route("/departments/<department>/files/new", methods=["GET", "POST"])
    @permission_required("add")
    def add_event_file(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        session["active_department"] = department
        events = Event.query.filter_by(department=department).order_by(Event.start_date.asc()).all()
        selected_event_id = optional_query_integer("event_id")
        if selected_event_id is not None and not any(event.id == selected_event_id for event in events):
            abort(404)
        form = make_event_file_form(events, selected_event_id=selected_event_id)
        goods_rows = event_goods_form_rows()
        parsed_goods, goods_errors = validate_event_goods_form_rows(goods_rows)
        file_form_valid = event_file_form_valid(form)
        if file_form_valid and not goods_errors:
            file_record = EventShipmentFile(created_by_user_id=current_user.id)
            apply_event_file_form(form, file_record)
            save_event_goods(file_record, parsed_goods)
            db.session.add(file_record)
            db.session.flush()
            assign_event_file_reference(
                file_record,
                file_transport_reference_code(form.transport_type.data, form.custom_transport.data),
            )
            snapshot = event_file_snapshot(file_record)
            audit_event(
                "create",
                "event_file",
                file_record.id,
                file_record.event.exhibition_name,
                {"fields": {key: {"before": None, "after": value} for key, value in snapshot.items()}},
            )
            db.session.commit()
            flash("Event file saved.", "success")
            return redirect(
                url_for(
                    "view_event_file",
                    department=department,
                    file_id=file_record.id,
                )
            )
        return render_template(
            "event_file_form.html",
            form=form,
            department=get_department(department),
            contacts_by_kind=saved_contacts_by_kind(),
            event_options=event_file_event_options(events),
            file_record=None,
            goods_rows=goods_rows,
            goods_errors=goods_errors,
            active_tab="goods" if goods_errors else "general",
            packing_types=PACKING_TYPES,
            clearance_methods=CLEARANCE_METHODS,
            page_title="Create event file",
            has_events=bool(events),
        )

    @app.get("/departments/<department>/files/<int:file_id>")
    @permission_required("view")
    def view_event_file(department, file_id):
        department_info = get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        session["active_department"] = department
        return render_template(
            "event_file_view.html",
            department=department_info,
            file_record=file_record,
        )

    @app.route("/departments/<department>/files/<int:file_id>/edit", methods=["GET", "POST"])
    @permission_required("edit")
    def edit_event_file(department, file_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        session["active_department"] = department
        events = Event.query.filter_by(department=department).order_by(Event.start_date.asc()).all()
        form = make_event_file_form(events, file_record=file_record)
        goods_rows = event_goods_form_rows(file_record)
        parsed_goods, goods_errors = validate_event_goods_form_rows(goods_rows)
        file_form_valid = event_file_form_valid(form)
        if file_form_valid and not goods_errors:
            before = event_file_snapshot(file_record)
            previous_event_id = file_record.event_id
            previous_transport_code = (
                file_record.reference_record.transport_code
                if file_record.reference_record
                else None
            )
            apply_event_file_form(form, file_record)
            save_event_goods(file_record, parsed_goods)
            transport_code = file_transport_reference_code(
                form.transport_type.data,
                form.custom_transport.data,
            )
            if (
                file_record.reference_record is None
                or previous_event_id != file_record.event_id
                or previous_transport_code != transport_code
            ):
                assign_event_file_reference(file_record, transport_code)
            after = event_file_snapshot(file_record)
            if before != after:
                audit_event(
                    "edit",
                    "event_file",
                    file_record.id,
                    file_record.event.exhibition_name,
                    {
                        "fields": {
                            key: {"before": before[key], "after": after[key]}
                            for key in before
                            if before[key] != after[key]
                        }
                    },
                )
                db.session.commit()
                flash("Event file updated.", "success")
            else:
                flash("No changes were made.", "warning")
            return redirect(
                url_for(
                    "view_event_file",
                    department=department,
                    file_id=file_record.id,
                )
            )
        return render_template(
            "event_file_form.html",
            form=form,
            department=get_department(department),
            contacts_by_kind=saved_contacts_by_kind(),
            event_options=event_file_event_options(events),
            file_record=file_record,
            goods_rows=goods_rows,
            goods_errors=goods_errors,
            active_tab="goods" if goods_errors else "general",
            packing_types=PACKING_TYPES,
            clearance_methods=CLEARANCE_METHODS,
            page_title="Edit event file",
            has_events=bool(events),
        )

    @app.get("/departments/<department>/files/<int:file_id>/proposals")
    @permission_required("view")
    def event_file_proposals(department, file_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        session["active_department"] = department

        def render_proposal_page(
            form_values,
            item_rows,
            errors=None,
            editing_proposal=None,
            form_action=None,
            submit_label="Save quotation",
        ):
            return render_template(
                "event_file_proposals.html",
                department=get_department(department),
                file_record=file_record,
                **event_file_proposal_template_data(file_record),
                proposals=file_record.proposals,
                editing_proposal=editing_proposal,
                form_values=form_values,
                item_rows=item_rows,
                errors=errors or [],
                currency_options=PROPOSAL_CURRENCY_OPTIONS,
                status_options=PROPOSAL_STATUS_OPTIONS,
                bank_accounts=event_file_proposal_bank_slots(),
                bank_form_action=url_for(
                    "save_proposal_bank_accounts",
                    department=department,
                    file_id=file_record.id,
                ),
                form_action=form_action
                or url_for(
                    "create_event_file_proposal",
                    department=department,
                    file_id=file_record.id,
                ),
                submit_label=submit_label,
                page_title="Proposal & quotation builder",
            )

        return render_proposal_page(
            event_file_proposal_form_defaults(file_record),
            event_file_proposal_item_rows(),
        )

    @app.route(
        "/departments/<department>/files/<int:file_id>/delivery-note",
        methods=["GET", "POST"],
    )
    @permission_required("view")
    def event_file_delivery_note(department, file_id):
        department_info = get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        session["active_department"] = department
        note_data = event_file_delivery_note_data(file_record)
        form_values = {
            "delivery_date": (
                file_record.delivery_note.delivery_date.isoformat()
                if file_record.delivery_note
                else date.today().isoformat()
            ),
            "condition_at_delivery": (
                file_record.delivery_note.condition_at_delivery
                if file_record.delivery_note
                else ""
            ),
            "receiver_name": file_record.delivery_note.receiver_name if file_record.delivery_note else "",
            "receiver_designation": (
                file_record.delivery_note.receiver_designation
                if file_record.delivery_note
                else ""
            ),
            "receiver_phone": file_record.delivery_note.receiver_phone if file_record.delivery_note else "",
            "receiver_email": file_record.delivery_note.receiver_email if file_record.delivery_note else "",
            "supervisor_name": file_record.delivery_note.supervisor_name if file_record.delivery_note else "",
            "supervisor_designation": (
                file_record.delivery_note.supervisor_designation
                if file_record.delivery_note
                else ""
            ),
            "supervisor_phone": (
                file_record.delivery_note.supervisor_phone
                if file_record.delivery_note
                else ""
            ),
        }
        errors = []
        can_save_delivery_note = (
            current_user.has_permission("edit")
            if file_record.delivery_note
            else current_user.has_permission("add") or current_user.has_permission("edit")
        )
        if request.method == "POST":
            action = request.form.get("action", "download")
            if action not in {"save", "download"}:
                abort(400)
            if action == "save" and not can_save_delivery_note:
                abort(403)
            form_values.update(
                {
                    field: request.form.get(field, "").strip()[:maximum_length]
                    for field, maximum_length in (
                        ("delivery_date", 10),
                        ("condition_at_delivery", 2000),
                        ("receiver_name", 200),
                        ("receiver_designation", 120),
                        ("receiver_phone", 50),
                        ("receiver_email", 254),
                        ("supervisor_name", 200),
                        ("supervisor_designation", 120),
                        ("supervisor_phone", 50),
                    )
                }
            )
            try:
                delivery_date = date.fromisoformat(form_values["delivery_date"])
            except ValueError:
                errors.append("Enter a valid delivery date.")
                delivery_date = None
            email = form_values["receiver_email"]
            if email and (
                "@" not in email or "." not in email.rsplit("@", 1)[-1]
            ):
                errors.append("Enter a valid receiver email address.")
            if not errors:
                if action == "save":
                    delivery_note = file_record.delivery_note
                    is_new = delivery_note is None
                    if is_new:
                        delivery_note = EventFileDeliveryNote(
                            file=file_record,
                            created_by_user_id=current_user.id,
                        )
                        db.session.add(delivery_note)
                    delivery_note_fields = (
                        "condition_at_delivery",
                        "receiver_name",
                        "receiver_designation",
                        "receiver_phone",
                        "receiver_email",
                        "supervisor_name",
                        "supervisor_designation",
                        "supervisor_phone",
                    )
                    before = {}
                    if not is_new:
                        before = {
                            "delivery_date": delivery_note.delivery_date.isoformat(),
                            **{
                                field: getattr(delivery_note, field)
                                for field in delivery_note_fields
                            },
                        }
                    delivery_note.delivery_date = delivery_date
                    for field in delivery_note_fields:
                        setattr(delivery_note, field, form_values[field])
                    after = {
                        "delivery_date": delivery_date.isoformat(),
                        **{
                            field: form_values[field]
                            for field in delivery_note_fields
                        },
                    }
                    if is_new:
                        audit_event(
                            "create",
                            "delivery_note",
                            file_record.id,
                            file_record.file_reference,
                            {"fields": after},
                        )
                    else:
                        if before != after:
                            audit_event(
                                "edit",
                                "delivery_note",
                                file_record.id,
                                file_record.file_reference,
                                {
                                    "fields": {
                                        field: {
                                            "before": before[field],
                                            "after": after[field],
                                        }
                                        for field in after
                                        if before[field] != after[field]
                                    }
                                },
                            )
                    file_record.updated_at = utc_now()
                    db.session.add(
                        EventFileDeliveryNoteHistory(
                            file_id=file_record.id,
                            delivery_date=delivery_date,
                            **{
                                field: form_values[field]
                                for field in delivery_note_fields
                            },
                            created_by_user_id=current_user.id,
                        )
                    )
                    db.session.commit()
                    flash("Delivery note saved.", "success")
                    return redirect(
                        url_for(
                            "event_file_delivery_note",
                            department=department,
                            file_id=file_record.id,
                        )
                    )
                output = build_event_file_delivery_note_pdf(
                    file_record,
                    note_data,
                    form_values,
                    delivery_date,
                )
                reference = secure_filename(file_record.file_reference or f"file-{file_record.id}")
                return send_file(
                    output,
                    mimetype="application/pdf",
                    as_attachment=True,
                    download_name=f"delivery-note-import-{reference}.pdf",
                )
        return render_template(
            "event_file_delivery_note.html",
            department=department_info,
            file_record=file_record,
            note_data=note_data,
            form_values=form_values,
            errors=errors,
            can_save_delivery_note=can_save_delivery_note,
            history=file_record.delivery_note_history,
        )

    @app.post("/departments/<department>/files/<int:file_id>/proposals")
    @permission_required("add")
    def create_event_file_proposal(department, file_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        session["active_department"] = department
        bank_accounts = ProposalBankAccount.query.all()
        form_state = event_file_proposal_request_state(request.form)
        try:
            values = parse_event_file_proposal_request(
                request.form,
                request.files,
                bank_accounts,
                form_state=form_state,
            )
        except ValueError as error:
            return render_template(
                "event_file_proposals.html",
                department=get_department(department),
                file_record=file_record,
                **event_file_proposal_template_data(file_record),
                proposals=file_record.proposals,
                editing_proposal=None,
                form_values=form_state[0],
                item_rows=form_state[1],
                errors=[str(error)],
                currency_options=PROPOSAL_CURRENCY_OPTIONS,
                status_options=PROPOSAL_STATUS_OPTIONS,
                bank_accounts=event_file_proposal_bank_slots(),
                bank_form_action=url_for(
                    "save_proposal_bank_accounts",
                    department=department,
                    file_id=file_record.id,
                ),
                form_action=url_for(
                    "create_event_file_proposal",
                    department=department,
                    file_id=file_record.id,
                ),
                submit_label="Save quotation",
                page_title="Proposal & quotation builder",
            )

        proposal = EventFileProposal(file=file_record, created_by_user_id=current_user.id)
        db.session.add(proposal)
        apply_event_file_proposal_values(proposal, values)
        db.session.flush()
        proposal.quote_number = f"RER-QTN-{utc_now():%Y%m%d}-{proposal.id:05d}"
        audit_event(
            "create",
            "event_file_proposal",
            proposal.id,
            proposal.quote_number,
            {
                "file_id": file_record.id,
                "currency": proposal.currency,
                "total": str(proposal.total),
                "status": proposal.status,
            },
        )
        db.session.commit()
        flash(f"Quotation {proposal.quote_number} created.", "success")
        return redirect(
            url_for("event_file_proposals", department=department, file_id=file_record.id)
        )

    @app.route(
        "/departments/<department>/files/<int:file_id>/proposals/<int:proposal_id>/edit",
        methods=["GET", "POST"],
    )
    @permission_required("edit")
    def edit_event_file_proposal(department, file_id, proposal_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        proposal = db.get_or_404(EventFileProposal, proposal_id)
        if file_record.event.department != department or proposal.file_id != file_record.id:
            abort(404)
        session["active_department"] = department
        form_action = url_for(
            "edit_event_file_proposal",
            department=department,
            file_id=file_record.id,
            proposal_id=proposal.id,
        )
        if request.method == "GET":
            return render_template(
                "event_file_proposals.html",
                department=get_department(department),
                file_record=file_record,
                **event_file_proposal_template_data(file_record),
                proposals=file_record.proposals,
                editing_proposal=proposal,
                form_values=event_file_proposal_form_defaults(file_record, proposal),
                item_rows=event_file_proposal_item_rows(proposal),
                errors=[],
                currency_options=PROPOSAL_CURRENCY_OPTIONS,
                status_options=PROPOSAL_STATUS_OPTIONS,
                bank_accounts=event_file_proposal_bank_slots(),
                bank_form_action=url_for(
                    "save_proposal_bank_accounts",
                    department=department,
                    file_id=file_record.id,
                ),
                form_action=form_action,
                submit_label="Update quotation",
                page_title=f"Edit {proposal.quote_number}",
            )

        bank_accounts = ProposalBankAccount.query.all()
        form_state = event_file_proposal_request_state(request.form)
        try:
            values = parse_event_file_proposal_request(
                request.form,
                request.files,
                bank_accounts,
                existing_proposal=proposal,
                form_state=form_state,
            )
        except ValueError as error:
            return render_template(
                "event_file_proposals.html",
                department=get_department(department),
                file_record=file_record,
                **event_file_proposal_template_data(file_record),
                proposals=file_record.proposals,
                editing_proposal=proposal,
                form_values=form_state[0],
                item_rows=form_state[1],
                errors=[str(error)],
                currency_options=PROPOSAL_CURRENCY_OPTIONS,
                status_options=PROPOSAL_STATUS_OPTIONS,
                bank_accounts=event_file_proposal_bank_slots(),
                bank_form_action=url_for(
                    "save_proposal_bank_accounts",
                    department=department,
                    file_id=file_record.id,
                ),
                form_action=form_action,
                submit_label="Update quotation",
                page_title=f"Edit {proposal.quote_number}",
            )

        apply_event_file_proposal_values(proposal, values)
        audit_event(
            "edit",
            "event_file_proposal",
            proposal.id,
            proposal.quote_number,
            {
                "file_id": file_record.id,
                "currency": proposal.currency,
                "total": str(proposal.total),
                "status": proposal.status,
            },
        )
        db.session.commit()
        flash(f"Quotation {proposal.quote_number} updated.", "success")
        return redirect(
            url_for("event_file_proposals", department=department, file_id=file_record.id)
        )

    @app.post("/departments/<department>/files/<int:file_id>/proposals/banks")
    @permission_required("edit")
    def save_proposal_bank_accounts(department, file_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        bank_fields = (
            ("label", 80),
            ("beneficiary", 200),
            ("bank_name", 200),
            ("account_number", 100),
            ("account_type", 40),
            ("currency", 3),
            ("branch", 200),
            ("ifsc", 30),
            ("swift", 20),
        )
        account_values = []
        for slot in (1, 2):
            values = {
                field: request.form.get(f"bank_{slot}_{field}", "").strip()
                for field, _ in bank_fields
            }
            for field, maximum_length in bank_fields:
                values[field] = values[field][:maximum_length]
            entered_account = any(
                values[field]
                for field in ("beneficiary", "bank_name", "account_number")
            )
            if entered_account and not all(
                values[field] for field in ("beneficiary", "bank_name", "account_number")
            ):
                flash(
                    f"Bank option {slot} needs a beneficiary, bank name and account number.",
                    "error",
                )
                return redirect(
                    url_for("event_file_proposals", department=department, file_id=file_id)
                )
            if values["currency"]:
                values["currency"] = values["currency"].upper()
                valid_currencies = {code for code, _ in PROPOSAL_CURRENCY_OPTIONS}
                if values["currency"] not in valid_currencies:
                    flash(f"Choose a valid currency for bank option {slot}.", "error")
                    return redirect(
                        url_for("event_file_proposals", department=department, file_id=file_id)
                    )
            account_values.append((slot, values))

        for slot, values in account_values:
            account = db.session.get(ProposalBankAccount, slot)
            if account is None:
                account = ProposalBankAccount(slot=slot)
                db.session.add(account)
            for field, _ in bank_fields:
                setattr(account, field, values[field])
            if not account.label:
                account.label = f"Bank option {slot}"
        db.session.commit()
        flash("Quotation bank options saved.", "success")
        return redirect(
            url_for("event_file_proposals", department=department, file_id=file_record.id)
        )

    @app.get(
        "/departments/<department>/files/<int:file_id>/proposals/<int:proposal_id>/pdf"
    )
    @permission_required("view")
    def download_event_file_proposal_pdf(department, file_id, proposal_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        proposal = db.get_or_404(EventFileProposal, proposal_id)
        if file_record.event.department != department or proposal.file_id != file_record.id:
            abort(404)
        output = build_event_file_proposal_pdf(proposal)
        return send_file(
            output,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{proposal.quote_number}.pdf",
        )

    @app.get(
        "/departments/<department>/files/<int:file_id>/proposals/<int:proposal_id>/upload"
    )
    @permission_required("view")
    def download_event_file_proposal_upload(department, file_id, proposal_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        proposal = db.get_or_404(EventFileProposal, proposal_id)
        if (
            file_record.event.department != department
            or proposal.file_id != file_record.id
            or not proposal.uploaded_data
        ):
            abort(404)
        return send_file(
            BytesIO(proposal.uploaded_data),
            mimetype=proposal.uploaded_mimetype,
            as_attachment=True,
            download_name=proposal.uploaded_filename,
        )

    @app.post("/departments/<department>/files/<int:file_id>/tracking-status")
    @permission_required("edit")
    def update_event_file_tracking_status(department, file_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department or file_record.transport_type != "Air":
            abort(404)

        new_status = request.form.get("tracking_status", "").strip()
        if new_status not in TRACKING_STATUSES:
            abort(400)

        air_details = file_record.air_details
        old_status = air_details.tracking_status if air_details else None
        if air_details is None:
            air_details = EventShipmentFileAirDetails(file=file_record)
            file_record.air_details = air_details
        if old_status != new_status:
            air_details.tracking_status = new_status
            file_record.updated_at = utc_now()
            audit_event(
                "status_updated",
                "event_file",
                file_record.id,
                file_record.event.exhibition_name,
                {
                    "fields": {
                        "tracking_status": {
                            "before": old_status,
                            "after": new_status,
                        }
                    }
                },
            )
            db.session.commit()
            flash("Event file tracking status updated.", "success")
        else:
            flash("No changes were made.", "warning")

        redirect_args = {"q": request.args.get("q", "").strip()}
        selected_event_id = optional_query_integer("event_id")
        if selected_event_id is not None:
            redirect_args["event_id"] = selected_event_id
        else:
            redirect_args["event_id"] = file_record.event_id
        return redirect(url_for("department_files", department=department, **redirect_args))

    @app.post("/departments/<department>/files/<int:file_id>/delete")
    @permission_required("delete")
    def delete_event_file(department, file_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        file_record = db.get_or_404(EventShipmentFile, file_id)
        if file_record.event.department != department:
            abort(404)
        event_id = file_record.event_id
        snapshot = event_file_snapshot(file_record)
        audit_event(
            "delete",
            "event_file",
            file_record.id,
            file_record.event.exhibition_name,
            {"fields": {key: {"before": value, "after": None} for key, value in snapshot.items()}},
        )
        db.session.delete(file_record)
        db.session.commit()
        flash("Event file deleted.", "success")
        return redirect(url_for("department_files", department=department, event_id=event_id))

    @app.post("/departments/<department>/airport-suggestions")
    @login_required
    def airport_location_suggestions(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        if not any(
            current_user.has_permission(permission)
            for permission in ("view", "add", "edit")
        ):
            abort(403)

        query = request.form.get("query", "").strip()
        if len(query) > 100:
            return jsonify(error="Search term is too long.", suggestions=[]), 400
        if len(query) < 2:
            return jsonify(suggestions=[])

        cache_path = current_app.config.get("AIRPORT_CACHE_PATH") or (
            Path(current_app.instance_path) / "ourairports.sqlite3"
        )
        try:
            suggestions = airport_suggestions(cache_path, query)
        except AirportLookupError:
            current_app.logger.exception("Could not load the local airport directory.")
            return jsonify(
                error="Airport suggestions are unavailable right now. You can enter the airport manually.",
                suggestions=[],
            ), 503

        for suggestion in suggestions:
            country = pycountry.countries.get(alpha_2=suggestion["country_code"])
            suggestion["country"] = country.name if country else suggestion["country_code"]
        return jsonify(suggestions=suggestions)

    @app.get("/departments/<department>/agents")
    @permission_required("view")
    def agents(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        session["active_department"] = department
        records = Contact.query.filter_by(kind="agent").order_by(Contact.company_name.asc()).all()
        return render_template("agents.html", department=get_department(department), agents=records)

    @app.get("/departments/<department>/agents.xlsx")
    @permission_required("view")
    def export_agents_excel(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        agents = Contact.query.filter_by(kind="agent").order_by(Contact.company_name.asc()).all()
        records = []
        for agent in agents:
            profile = agent.agent_profile
            people = (
                "\n".join(
                    f"{person.contact_name}: {person.phone}"
                    for person in profile.contact_people
                )
                if profile and profile.contact_people
                else f"{agent.contact_person_name}: {agent.phone}".strip(": ")
            )
            records.append(
                SimpleNamespace(
                    company_name=agent.company_name,
                    country=profile.country if profile else "",
                    address=profile.address if profile else "",
                    city=profile.city if profile else "",
                    postal_code=profile.postal_code if profile else "",
                    company_phone=(
                        profile.company_phone if profile and profile.company_phone else agent.phone
                    ),
                    company_email=(
                        profile.company_email if profile and profile.company_email else agent.email
                    ),
                    vat_number=profile.vat_number if profile else "",
                    contact_people=people,
                    updated_at=agent.updated_at,
                )
            )
        output = create_excel_export(
            records,
            fields=AGENT_EXPORT_FIELDS,
            sheet_title="Agents",
        )
        filename = f"agents_{utc_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    @app.get("/departments/<department>/customers")
    @permission_required("view")
    def customers(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        session["active_department"] = department
        records = Contact.query.filter_by(kind="customer").order_by(Contact.company_name.asc()).all()
        return render_template("customers.html", department=get_department(department), customers=records)

    @app.get("/departments/<department>/customers.xlsx")
    @permission_required("view")
    def export_customers_excel(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        records = Contact.query.filter_by(kind="customer").order_by(Contact.company_name.asc()).all()
        output = create_excel_export(
            records,
            fields=CUSTOMER_EXPORT_FIELDS,
            sheet_title="Customers",
        )
        filename = f"customers_{utc_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    @app.route("/departments/<department>/customers/<int:contact_id>/edit", methods=["GET", "POST"])
    @permission_required("edit")
    def edit_customer(department, contact_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        customer = db.get_or_404(Contact, contact_id)
        if customer.kind != "customer":
            abort(404)
        session["active_department"] = department
        form = CustomerForm()
        if request.method == "GET":
            form.company_name.data = customer.company_name
            form.contact_person_name.data = customer.contact_person_name
            form.address.data = customer.address
            form.email.data = customer.email
            form.phone.data = customer.phone
        if form.validate_on_submit():
            company_name = form.company_name.data.strip()
            name_key = company_name.casefold()
            duplicate = Contact.query.filter_by(kind="customer", name_key=name_key).first()
            if duplicate and duplicate.id != customer.id:
                form.company_name.errors.append(
                    "A customer with this name already exists in the directory."
                )
            else:
                before = contact_snapshot(customer)
                customer.company_name = company_name
                customer.name_key = name_key
                customer.contact_person_name = (form.contact_person_name.data or "").strip()
                customer.address = (form.address.data or "").strip()
                customer.email = (form.email.data or "").strip()
                customer.phone = (form.phone.data or "").strip()
                after = contact_snapshot(customer)
                if before != after:
                    customer.updated_at = utc_now()
                    audit_event(
                        "edit",
                        "contact",
                        customer.id,
                        customer.company_name,
                        {
                            "fields": {
                                field: {"before": before[field], "after": after[field]}
                                for field in before
                                if before[field] != after[field]
                            }
                        },
                    )
                    db.session.commit()
                    flash("Customer details updated.", "success")
                else:
                    flash("No changes were made.", "warning")
                return redirect(url_for("customers", department=department))
        return render_template(
            "customer_form.html",
            form=form,
            department=get_department(department),
            page_title="Edit customer",
        )

    @app.route("/departments/<department>/agents/new", methods=["GET", "POST"])
    @permission_required("add")
    def add_agent(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        session["active_department"] = department
        form = AgentProfileForm()
        if form.validate_on_submit():
            save_agent_profile(form)
            db.session.commit()
            flash("Agent details saved to the directory.", "success")
            return redirect(url_for("agents", department=department))
        return render_template(
            "agent_form.html",
            form=form,
            department=get_department(department),
            countries=COUNTRY_OPTIONS,
            page_title="Add agent",
        )

    @app.route("/departments/<department>/agents/<int:contact_id>/edit", methods=["GET", "POST"])
    @permission_required("edit")
    def edit_agent(department, contact_id):
        get_department(department)
        if department != "import-export":
            abort(404)
        contact = db.get_or_404(Contact, contact_id)
        if contact.kind != "agent":
            abort(404)
        session["active_department"] = department
        form = AgentProfileForm()
        if request.method == "GET":
            populate_agent_profile_form(form, contact)
        if form.validate_on_submit():
            new_name = form.company_name.data.strip()
            new_key = new_name.casefold()
            duplicate = Contact.query.filter_by(kind="agent", name_key=new_key).first()
            if duplicate and duplicate.id != contact.id:
                form.company_name.errors.append("An agent with this name already exists in the directory.")
            else:
                save_agent_profile(form, contact)
                db.session.commit()
                flash("Agent details updated.", "success")
                return redirect(url_for("agents", department=department))
        return render_template(
            "agent_form.html",
            form=form,
            department=get_department(department),
            countries=COUNTRY_OPTIONS,
            page_title="Edit agent",
        )

    @app.post("/departments/<department>/agent-city-suggestions")
    @login_required
    def agent_city_suggestions(department):
        get_department(department)
        if department != "import-export":
            abort(404)
        if not any(
            current_user.has_permission(permission)
            for permission in ("view", "add", "edit")
        ):
            abort(403)

        country_code = request.form.get("country_code", "").strip().upper()
        if not pycountry.countries.get(alpha_2=country_code):
            return jsonify(error="Select a country from the suggestions.", suggestions=[]), 400
        postal_code = request.form.get("postal_code", "").strip()
        if len("".join(character for character in postal_code if character.isalnum())) < 2:
            return jsonify(available=True, suggestions=[])

        cache_path = current_app.config.get("POSTAL_CODE_CACHE_PATH") or (
            Path(current_app.instance_path) / "geonames-postal-codes.sqlite3"
        )
        try:
            result = postal_code_suggestions(cache_path, country_code, postal_code)
        except PostalCodeLookupError:
            current_app.logger.exception(
                "Could not load local postal-code data for country %s",
                country_code,
            )
            return jsonify(
                error="Postal-code suggestions are unavailable right now. You can enter the city manually.",
                suggestions=[],
            ), 503
        return jsonify(result)

    @app.get("/shipments")
    @permission_required("view")
    def shipments():
        search = request.args.get("q", "").strip()
        status_filter = request.args.get("status", "").strip()
        if status_filter and status_filter not in TRACKING_STATUSES:
            abort(400)

        records = filtered_shipment_query(search, status_filter).all()
        status_counts = {
            status: Shipment.query.filter_by(tracking_status=status).count()
            for status in TRACKING_STATUSES
        }
        return render_template(
            "shipments.html",
            shipments=records,
            search=search,
            status_filter=status_filter,
            status_counts=status_counts,
            total_count=Shipment.query.count(),
            tracking_statuses=TRACKING_STATUSES,
        )

    @app.get("/exports/shipments.xlsx")
    @permission_required("view")
    def export_excel():
        search = request.args.get("q", "").strip()
        status_filter = request.args.get("status", "").strip()
        if status_filter and status_filter not in TRACKING_STATUSES:
            abort(400)

        output = create_excel_export(filtered_shipment_query(search, status_filter).all())
        filename = f"shipments_{utc_now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    @app.get("/exports/shipments.pdf")
    @permission_required("view")
    def export_pdf():
        search = request.args.get("q", "").strip()
        status_filter = request.args.get("status", "").strip()
        if status_filter and status_filter not in TRACKING_STATUSES:
            abort(400)

        records = filtered_shipment_query(search, status_filter).all()
        output = create_pdf_export(records, search, status_filter)
        filename = f"shipments_{utc_now().strftime('%Y%m%d_%H%M%S')}.pdf"
        return send_file(
            output,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    @app.route("/shipments/new", methods=["GET", "POST"])
    @permission_required("add")
    def add_shipment():
        form = make_shipment_form()
        if form.validate_on_submit():
            reference = form.reference.data.strip()
            reference_key = reference.casefold()
            if Shipment.query.filter_by(reference_key=reference_key).first():
                form.reference.errors.append("This shipment reference already exists.")
            else:
                shipment = Shipment(reference=reference, reference_key=reference_key)
                form.populate_obj(shipment)
                shipment.reference = reference
                shipment.reference_key = reference_key
                db.session.add(shipment)
                db.session.flush()
                snapshot = shipment_snapshot(shipment)
                audit_event(
                    "create",
                    "shipment",
                    shipment.id,
                    reference,
                    {"fields": {key: {"before": None, "after": value} for key, value in snapshot.items()}},
                )
                db.session.commit()
                flash("Shipment saved.", "success")
                destination = "shipments" if current_user.has_permission("view") else "add_shipment"
                return redirect(url_for(destination))
        return render_template("shipment_form.html", form=form, page_title="New shipment")

    @app.route("/shipments/<int:shipment_id>/edit", methods=["GET", "POST"])
    @permission_required("edit")
    def edit_shipment(shipment_id):
        shipment = db.get_or_404(Shipment, shipment_id)
        form = make_shipment_form(shipment)
        if form.validate_on_submit():
            reference = form.reference.data.strip()
            reference_key = reference.casefold()
            duplicate = Shipment.query.filter_by(reference_key=reference_key).first()
            if duplicate and duplicate.id != shipment.id:
                form.reference.errors.append("This shipment reference already exists.")
            else:
                before = shipment_snapshot(shipment)
                form.populate_obj(shipment)
                shipment.reference = reference
                shipment.reference_key = reference_key
                after = shipment_snapshot(shipment)
                changes = {
                    key: {"before": before[key], "after": after[key]}
                    for key in before
                    if before[key] != after[key]
                }
                if changes:
                    shipment.updated_at = utc_now()
                    audit_event(
                        "edit",
                        "shipment",
                        shipment.id,
                        reference,
                        {"fields": changes},
                    )
                    db.session.commit()
                    flash("Shipment updated.", "success")
                else:
                    flash("No changes were made.", "warning")
                return redirect(url_for("shipments"))
        return render_template("shipment_form.html", form=form, page_title="Shipment update")

    @app.post("/shipments/<int:shipment_id>/tracking-status")
    @permission_required("edit")
    def update_tracking_status(shipment_id):
        shipment = db.get_or_404(Shipment, shipment_id)
        new_status = request.form.get("tracking_status", "").strip()
        if new_status not in TRACKING_STATUSES:
            abort(400)

        old_status = shipment.tracking_status
        if old_status != new_status:
            shipment.tracking_status = new_status
            shipment.updated_at = utc_now()
            audit_event(
                "status_updated",
                "shipment",
                shipment.id,
                shipment.reference,
                {
                    "fields": {
                        "tracking_status": {
                            "before": old_status,
                            "after": new_status,
                        }
                    }
                },
            )
            db.session.commit()
            flash(f"{shipment.reference} status updated to {new_status}.", "success")
        else:
            flash("No status changes were made.", "warning")

        return redirect(
            url_for(
                "shipments",
                q=request.args.get("q", ""),
                status=request.args.get("status", ""),
            )
        )

    @app.post("/shipments/<int:shipment_id>/delete")
    @permission_required("delete")
    def delete_shipment(shipment_id):
        shipment = db.get_or_404(Shipment, shipment_id)
        snapshot = shipment_snapshot(shipment)
        audit_event(
            "delete",
            "shipment",
            shipment.id,
            shipment.reference,
            {"fields": {key: {"before": value, "after": None} for key, value in snapshot.items()}},
        )
        db.session.delete(shipment)
        db.session.commit()
        flash("Shipment deleted.", "success")
        return redirect(url_for("shipments"))

    @app.route("/users", methods=["GET", "POST"])
    @admin_required
    def users():
        form = CreateUserForm()
        if form.validate_on_submit():
            username = form.username.data.strip()
            if User.query.filter_by(username=username).first():
                form.username.errors.append("This username is already in use.")
            else:
                user = User(
                    username=username,
                    is_admin=False,
                    can_view=(
                        form.can_view.data
                        or form.can_add.data
                        or form.can_edit.data
                        or form.can_delete.data
                    ),
                    can_add=form.can_add.data,
                    can_edit=form.can_edit.data,
                    can_delete=form.can_delete.data,
                )
                user.set_password(form.password.data)
                db.session.add(user)
                db.session.flush()
                audit_event(
                    "create",
                    "user",
                    user.id,
                    username,
                    {
                        "permissions": {
                            "view": user.can_view,
                            "add": user.can_add,
                            "edit": user.can_edit,
                            "delete": user.can_delete,
                        }
                    },
                )
                db.session.commit()
                flash(f"Login account created for {username}.", "success")
                return redirect(url_for("users"))
        user_list = User.query.order_by(User.username.asc()).all()
        return render_template("users.html", form=form, users=user_list)

    @app.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
    @admin_required
    def edit_user(user_id):
        user = db.get_or_404(User, user_id)
        form = ManageUserForm(obj=user)
        if form.validate_on_submit():
            if not user.is_admin:
                before_permissions = {
                    "view": user.can_view,
                    "add": user.can_add,
                    "edit": user.can_edit,
                    "delete": user.can_delete,
                }
                user.can_view = (
                    form.can_view.data
                    or form.can_add.data
                    or form.can_edit.data
                    or form.can_delete.data
                )
                user.can_add = form.can_add.data
                user.can_edit = form.can_edit.data
                user.can_delete = form.can_delete.data
                after_permissions = {
                    "view": user.can_view,
                    "add": user.can_add,
                    "edit": user.can_edit,
                    "delete": user.can_delete,
                }
                if before_permissions != after_permissions:
                    audit_event(
                        "permissions_updated",
                        "user",
                        user.id,
                        user.username,
                        {"before": before_permissions, "after": after_permissions},
                    )

            password_reset = bool(form.new_password.data)
            if password_reset:
                user.set_password(form.new_password.data)
                active_sessions = UserSession.query.filter_by(user_id=user.id, ended_at=None).all()
                for activity_session in active_sessions:
                    activity_session.ended_at = utc_now()
                    activity_session.end_reason = "password_reset"
                audit_event("password_reset", "user", user.id, user.username, {})

            db.session.commit()
            flash(f"Access settings saved for {user.username}.", "success")
            return redirect(url_for("users"))
        return render_template("user_edit.html", form=form, user=user)

    @app.get("/admin/activity")
    @admin_required
    def admin_activity():
        now = utc_now()
        close_idle_sessions(now)
        records = UserSession.query.order_by(UserSession.started_at.desc()).limit(200).all()
        activity_rows = []
        active_user_ids = set()
        for activity_session in records:
            is_active = (
                activity_session.ended_at is None
                and now - as_utc(activity_session.last_seen_at) <= ACTIVE_WINDOW
            )
            if is_active:
                active_user_ids.add(activity_session.user_id)
            activity_rows.append(
                {
                    "session": activity_session,
                    "is_active": is_active,
                    "duration": session_duration(activity_session, now),
                }
            )
        return render_template(
            "admin_activity.html",
            rows=activity_rows,
            active_user_count=len(active_user_ids),
        )

    @app.get("/admin/audit")
    @admin_required
    def admin_audit():
        records = AuditLog.query.order_by(AuditLog.created_at.desc()).limit(300).all()
        for record in records:
            record.details = json.loads(record.details_json)
        return render_template("admin_audit.html", records=records)

    @app.cli.command("create-admin")
    def create_admin():
        username = click.prompt("Admin username").strip()
        if len(username) < 3 or len(username) > 80:
            raise click.ClickException("Username must be between 3 and 80 characters.")
        if User.query.filter_by(username=username).first():
            raise click.ClickException("This username already exists.")
        password = click.prompt("Admin password (minimum 12 characters)", hide_input=True, confirmation_prompt=True)
        if len(password) < 12:
            raise click.ClickException("Password must be at least 12 characters.")
        user = User(username=username, is_admin=True)
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        click.echo("Admin account created.")

    return app

app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
