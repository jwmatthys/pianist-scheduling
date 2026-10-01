PRAGMA user_version = 0;

CREATE TABLE organizations (
    id INTEGER NOT NULL PRIMARY KEY,
    name VARCHAR(200) NOT NULL,
    created_at DATETIME NOT NULL
);

CREATE TABLE pianists (
    id INTEGER NOT NULL PRIMARY KEY,
    organization_id INTEGER NOT NULL,
    name VARCHAR(200) NOT NULL,
    email VARCHAR(200) NOT NULL,
    max_hours_per_week FLOAT,
    FOREIGN KEY(organization_id) REFERENCES organizations (id)
);

CREATE TABLE availability_slots (
    id INTEGER NOT NULL PRIMARY KEY,
    pianist_id INTEGER NOT NULL,
    day VARCHAR(20) NOT NULL,
    slot_start_minute INTEGER NOT NULL,
    status VARCHAR(20) NOT NULL,
    CONSTRAINT uq_slot UNIQUE (pianist_id, day, slot_start_minute),
    FOREIGN KEY(pianist_id) REFERENCES pianists (id)
);

CREATE TABLE lessons (
    id INTEGER NOT NULL PRIMARY KEY,
    organization_id INTEGER NOT NULL,
    teacher VARCHAR(200) NOT NULL,
    student VARCHAR(200) NOT NULL,
    day VARCHAR(20) NOT NULL,
    start_minute INTEGER NOT NULL,
    end_minute INTEGER NOT NULL,
    room VARCHAR(200) NOT NULL,
    instrument VARCHAR(200) NOT NULL,
    required_pianist_name VARCHAR(200) NOT NULL,
    need_pianist BOOLEAN NOT NULL,
    assigned_pianist_id INTEGER,
    fit_quality VARCHAR(20) NOT NULL,
    notes TEXT NOT NULL,
    hours FLOAT NOT NULL,
    manually_edited BOOLEAN NOT NULL,
    FOREIGN KEY(organization_id) REFERENCES organizations (id),
    FOREIGN KEY(assigned_pianist_id) REFERENCES pianists (id)
);

CREATE TABLE import_profiles (
    id INTEGER NOT NULL PRIMARY KEY,
    organization_id INTEGER NOT NULL,
    name VARCHAR(200) NOT NULL,
    mapping_json TEXT NOT NULL,
    FOREIGN KEY(organization_id) REFERENCES organizations (id)
);

INSERT INTO organizations (id, name, created_at)
VALUES (1, 'Synthetic Music Program', '2026-01-01 12:00:00');

INSERT INTO pianists (id, organization_id, name, email, max_hours_per_week)
VALUES (7, 1, 'Synthetic Pianist', 'pianist@example.invalid', 8.0);

INSERT INTO availability_slots (id, pianist_id, day, slot_start_minute, status)
VALUES (9, 7, 'Monday', 540, 'Available');

INSERT INTO lessons (
    id, organization_id, teacher, student, day, start_minute, end_minute,
    room, instrument, required_pianist_name, need_pianist,
    assigned_pianist_id, fit_quality, notes, hours, manually_edited
)
VALUES (
    11, 1, 'Synthetic Instructor', 'Synthetic Student', 'Monday', 540, 590,
    'Room 101', 'Violin', '', 1, 7, 'Full', '', 0.83, 0
);

INSERT INTO import_profiles (id, organization_id, name, mapping_json)
VALUES (13, 1, 'Synthetic Lesson Map', '{"student":"Student Name"}');