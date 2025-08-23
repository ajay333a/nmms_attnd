CREATE TABLE nmms_attndc (
    attnd_id INTEGER PRIMARY KEY AUTOINCREMENT,
    taluk_name TEXT,
    panchayath_name TEXT,
    work_code TEXT,
    nmr_number INT NOT NULL,
    job_card_num TEXT NOT NULL,
    worker_name TEXT,
    gender VARCHAR(1) NOT NULL CHECK(gender IN ('M', 'F', 'O')),
    attndc INTEGER NOT NULL CHECK(attndc IN (0,1)),
    attndnc_date DATETIME,
    photo_taker TEXT
);