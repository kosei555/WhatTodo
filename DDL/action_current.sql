CREATE TABLE action_current(
    id AUTOINCREMENT PRIMARY KEY,
    addedDate Date NOT NULL,
    todo LongText NOT NULL,
    done Bit NOT NULL,
    project_id Long
)