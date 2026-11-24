"""Persistence for application-owned license activation state."""

import json
from datetime import datetime
from pathlib import Path

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

from .. import database
from ..license_schemas import LicensePayload, LicenseRecord


def _create_engine(path: Path) -> Engine:
    path.parent.mkdir(parents=True, exist_ok=True)
    return create_engine(f"sqlite:///{path}", connect_args={"check_same_thread": False})


license_engine = _create_engine(database.LICENSE_DB_PATH)


class LicenseRepository:
    def __init__(self, engine: Engine):
        self._engine = engine

    def initialize(self) -> None:
        database.migrate_application_database(self._engine)

    def get_license_record(self) -> LicenseRecord | None:
        with self._engine.connect() as connection:
            row = connection.execute(
                text(
                    "SELECT payload_json, signature, activated_at "
                    "FROM license_activations WHERE singleton = 1"
                )
            ).mappings().one_or_none()
        if row is None:
            return None
        try:
            payload = LicensePayload.model_validate(json.loads(row["payload_json"]))
            activated_at = datetime.fromisoformat(row["activated_at"].replace("Z", "+00:00"))
            return LicenseRecord(
                payload=payload,
                signature=row["signature"],
                activated_at=activated_at,
            )
        except (ValueError, TypeError) as error:
            raise RuntimeError("Stored license activation data is malformed.") from error

    def save_license_record(self, payload: LicensePayload, signature: str, activated_at: datetime) -> LicenseRecord:
        payload_json = json.dumps(
            payload.model_dump(mode="json"),
            ensure_ascii=False,
            separators=(",", ":"),
            sort_keys=True,
        )
        timestamp = activated_at.isoformat(timespec="seconds").replace("+00:00", "Z")
        with self._engine.begin() as connection:
            connection.execute(
                text(
                    "INSERT INTO license_activations "
                    "(singleton, payload_json, signature, activated_at) "
                    "VALUES (1, :payload_json, :signature, :activated_at) "
                    "ON CONFLICT(singleton) DO UPDATE SET "
                    "payload_json = excluded.payload_json, "
                    "signature = excluded.signature, "
                    "activated_at = excluded.activated_at"
                ),
                {
                    "payload_json": payload_json,
                    "signature": signature,
                    "activated_at": timestamp,
                },
            )
        return LicenseRecord(payload=payload, signature=signature, activated_at=activated_at)

    def clear_license(self) -> None:
        with self._engine.begin() as connection:
            connection.execute(text("DELETE FROM license_activations WHERE singleton = 1"))


repository = LicenseRepository(license_engine)
