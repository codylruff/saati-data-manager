-- user information table
CREATE TABLE IF NOT EXISTS user_info (
    User_Id                 TEXT PRIMARY KEY,
    Password                TEXT,
    Preferences_Json        TEXT           
)