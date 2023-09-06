import sqlite3
from sqlite3 import Error

DATABASE_NAME = "email_subscribers.db"


def create_connection():
    """Create a database connection and return the connection object."""
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        return conn
    except Error as e:
        print(e)

    return conn


# This is just for testing purposes to ensure our connection works
if __name__ == "__main__":
    conn = create_connection()
    if conn:
        print("Successfully connected to the database!")
        conn.close()


def create_subscribers_table(conn):
    """Create the subscribers table if it doesn't exist."""
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS subscribers (
                id INTEGER PRIMARY KEY,
                email TEXT NOT NULL UNIQUE,
                is_subscribed INTEGER NOT NULL DEFAULT 1
            );
        """
        )
    except Error as e:
        print(e)


def add_email(conn, email):
    """Add a new email to the subscribers table."""
    try:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO subscribers(email) VALUES(?)", (email,))
        conn.commit()
        return cursor.lastrowid
    except Error as e:
        print(e)
        return None


def unsubscribe_email(conn, email):
    """Unsubscribe an email."""
    try:
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE subscribers SET is_subscribed = 0 WHERE email = ?", (email,)
        )
        conn.commit()
    except Error as e:
        print(e)


###TESTING###

if __name__ == "__main__":
    conn = create_connection()

    if conn:
        # Create the subscribers table
        create_subscribers_table(conn)

        # Test adding an email
        add_email(conn, "test@example.com")

        # Test unsubscribing the same email
        unsubscribe_email(conn, "test@example.com")

        print("Database operations completed!")
        conn.close()
