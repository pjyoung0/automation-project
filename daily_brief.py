import os

EMAIL = os.environ["EMAIL_ADDRESS"]
PASSWORD = os.environ["EMAIL_PASSWORD"]

print("EMAIL:", EMAIL)
print("PASSWORD LENGTH:", len(PASSWORD))
