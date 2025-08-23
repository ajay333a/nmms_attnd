CREATE TABLE nmms_attndnc (
    attndc_id INTEGER PRIMARY KEY AUTOINCREMENT,
    block_panch_id INT NOT NULL,
    work_id INT NOT NULL,
    msr_no INT NOT NULL,
    job_card_id INT NOT NULL,
    worker_id INT NOT NULL,
    gender VARCHAR(1) NOT NULL CHECK("M", "F", "O"),
    attndnc VARCHAR(1) NOT NULL CHECK("A", "B"),
    attnd_date DATETIME NOT NULL,
    photo_taker_id INT NOT NULL,
    FOREIGN KEY (loc_id) REFERENCES 
);

CREATE TABLE block_panch (
    block_panch_id INTEGER PRIMARY KEY AUTOINCREMENT,
    block_name TEXT NOT NULL,
    panchayath_name TEXT NOT NULL
);

CREATE TABLE works (
    work_id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_code TEXT NOT NULL,
    work_name TEXT NOT NULL,
    person_days INT ,
    achived_person_days INT,
    total_amount INT,
    wage_amount INT,
    material_amount INT
);


