CREATE TABLE IF NOT EXISTS det_data (
    ord_det NCHAR   NOT NULL PRIMARY KEY,
    ord     NCHAR   NOT NULL,
    det     NCHAR   NOT NULL,
    qty     INT     NOT NULL
);


CREATE TABLE IF NOT EXISTS calc_files (
    ord_det   NCHAR NOT NULL PRIMARY KEY,
    file_name NCHAR NOT NULL,
    FOREIGN KEY (ord_det) REFERENCES det_data (ord_det)
);


CREATE TABLE IF NOT EXISTS calculation (
    ord_det     NCHAR   NOT NULL,
    part_num    NCHAR   NOT NULL,
    part_qty    INT     NOT NULL,
    material    NCHAR   NOT NULL,
    mat_type    NCHAR   NOT NULL,
    mat_per_1   INT     NOT NULL,
    unit        NCHAR   NOT NULL,
    qty_per_det INT     NOT NULL,
    FOREIGN KEY (ord_det) REFERENCES det_data (ord_det),
    UNIQUE (ord_det, part_num)
);


CREATE TABLE IF NOT EXISTS halffabricat (
    ord_det  NCHAR NOT NULL,
    part_num NCHAR NOT NULL,
    qty      INT   NOT NULL,
    UNIQUE (ord_det, part_num)
);


CREATE TABLE IF NOT EXISTS hf_files (
    ord       NCHAR NOT NULL PRIMARY KEY,
    file_name NCHAR NOT NULL,
    FOREIGN KEY (ord) REFERENCES det_data (ord)
);
