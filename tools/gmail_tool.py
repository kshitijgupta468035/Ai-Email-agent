# ============================================================
# Gmail Tool
# ============================================================
# Purpose:
#     This file contains Gmail-related functionality for our
#     AI Email Agent.
#
# We will gradually add functions for:
#
#     1. Gmail authentication
#     2. Creating email messages
#     3. Sending emails
#
# Important:
#     We use Google's OAuth system instead of storing a Gmail
#     username and password in our Python code.
# ============================================================


# ------------------------------------------------------------
# Gmail API Scope
# ------------------------------------------------------------
# This scope gives our application permission to SEND emails
# through Gmail.
#
# We intentionally request only the permission that our
# current application needs.
# ------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send"
]

# ============================================================
# Import required Gmail / Google authentication libraries
# ============================================================

# os allows us to check whether files such as token.json
# already exist.
import os

# Credentials allows us to load previously saved Google
# authorization credentials from token.json.
from google.oauth2.credentials import Credentials

# InstalledAppFlow starts the OAuth authorization process
# when the user needs to give permission to the application.
from google_auth_oauthlib.flow import InstalledAppFlow

# build creates a Gmail API service object that our Python
# program can use to communicate with Gmail.
from googleapiclient.discovery import build

# ------------------------------------------------------------
# Email message libraries
# ------------------------------------------------------------

# MIMEText helps us create a properly formatted plain-text
# email message.
from email.mime.text import MIMEText

# base64 is required by the Gmail API to encode the email
# before sending it.
import base64

# ============================================================
# Function: authenticate_gmail
# ============================================================
# Purpose:
#     Authenticate our application with Gmail and return a
#     Gmail API service object.
#
# Authentication flow:
#
#     1. Check whether token.json already exists.
#     2. If it exists, load the saved credentials.
#     3. Check whether those credentials are valid.
#     4. If valid, reuse them.
#     5. If not valid, start Google's OAuth flow.
#     6. Save the new authorization information to token.json.
#     7. Create and return the Gmail API service.
#
# Required files:
#
#     credentials.json
#         Identifies our Google OAuth application.
#
#     token.json
#         Stores authorization obtained after the user
#         completes the OAuth process.
#
# Returns:
#
#     Gmail API service object.
#
# Raises:
#
#     FileNotFoundError:
#         If credentials.json is missing.
#
#     RuntimeError:
#         If Gmail authentication fails.
# ============================================================

def authenticate_gmail():

    try:

        # ----------------------------------------------------
        # Variable for storing Google credentials
        # ----------------------------------------------------
        # Initially we don't have credentials loaded.
        creds = None


        # ----------------------------------------------------
        # Check whether token.json already exists
        # ----------------------------------------------------
        # token.json contains authorization information from
        # a previous successful OAuth login.
        if os.path.exists("token.json"):

            # Load the saved credentials.
            creds = Credentials.from_authorized_user_file(
                "token.json",
                SCOPES
            )


        # ----------------------------------------------------
        # Check whether existing credentials are usable
        # ----------------------------------------------------
        # If credentials don't exist or are no longer valid,
        # we need to authenticate again.
        if not creds or not creds.valid:

            # ------------------------------------------------
            # Check that credentials.json exists
            # ------------------------------------------------
            # Without this file we cannot start the OAuth flow.
            if not os.path.exists("credentials.json"):

                raise FileNotFoundError(
                    "credentials.json was not found. "
                    "Please place your Google OAuth credentials "
                    "file in the project folder."
                )


            # ------------------------------------------------
            # Create the OAuth flow
            # ------------------------------------------------
            # credentials.json tells Google which OAuth
            # application is requesting access.
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json",
                SCOPES
            )


            # ------------------------------------------------
            # Ask the user to authorize the application
            # ------------------------------------------------
            # Google will open a browser window where the user
            # can sign in and approve the requested Gmail scope.
            creds = flow.run_local_server(port=0)


            # ------------------------------------------------
            # Save the authorization information
            # ------------------------------------------------
            # This prevents us from having to repeat the
            # authorization process every time the program runs.
            with open("token.json", "w") as token:

                token.write(
                    creds.to_json()
                )


        # ----------------------------------------------------
        # Create Gmail API service
        # ----------------------------------------------------
        # The service object is what we'll later use to perform
        # Gmail operations such as sending messages.
        gmail = build(
            "gmail",
            "v1",
            credentials=creds
        )


        # ----------------------------------------------------
        # Return Gmail service
        # ----------------------------------------------------
        return gmail


    except FileNotFoundError:

        # ----------------------------------------------------
        # Re-raise file-related errors
        # ----------------------------------------------------
        # The error message is already useful, so we don't
        # replace it.
        raise


    except Exception as e:

        # ----------------------------------------------------
        # Handle unexpected authentication errors
        # ----------------------------------------------------
        # Add useful context before passing the error upward.
        raise RuntimeError(
            f"Gmail authentication failed: {e}"
        ) from e


# ============================================================
# Function: send_email
# ============================================================
# Purpose:
#     Send a plain-text email through the Gmail API.
#
# Parameters:
#     gmail:
#         Gmail API service object returned by
#         authenticate_gmail().
#
#     recipient:
#         Email address of the person receiving the email.
#
#     subject:
#         Subject line of the email.
#
#     body:
#         Main text/content of the email.
#
# Returns:
#     dict:
#         Response returned by the Gmail API after successfully
#         sending the message.
#
# Raises:
#     ValueError:
#         If recipient, subject, or body is empty.
#
#     RuntimeError:
#         If Gmail fails to create or send the message.
# ============================================================

def send_email(gmail, recipient, subject, body):

    try:

        # ----------------------------------------------------
        # Validate recipient
        # ----------------------------------------------------
        # We don't want to send an email if there is no
        # recipient address.
        if not recipient or not str(recipient).strip():

            raise ValueError(
                "Recipient email address cannot be empty."
            )


        # ----------------------------------------------------
        # Validate subject
        # ----------------------------------------------------
        # The subject should contain some meaningful text.
        if not subject or not str(subject).strip():

            raise ValueError(
                "Email subject cannot be empty."
            )


        # ----------------------------------------------------
        # Validate email body
        # ----------------------------------------------------
        # We should never send a completely empty email.
        if not body or not str(body).strip():

            raise ValueError(
                "Email body cannot be empty."
            )


        # ----------------------------------------------------
        # Create the email message
        # ----------------------------------------------------
        # MIMEText creates a standard plain-text email.
        message = MIMEText(
            str(body),
            "plain"
        )


        # ----------------------------------------------------
        # Add email headers
        # ----------------------------------------------------
        # These become the visible To and Subject fields
        # of the email.
        message["To"] = str(recipient).strip()
        message["Subject"] = str(subject).strip()


        # ----------------------------------------------------
        # Encode the email
        # ----------------------------------------------------
        # Gmail API expects the raw email to be Base64URL
        # encoded.
        raw_message = base64.urlsafe_b64encode(
            message.as_bytes()
        ).decode()


        # ----------------------------------------------------
        # Create Gmail API request
        # ----------------------------------------------------
        # userId="me" means the currently authenticated
        # Gmail account.
        request_body = {
            "raw": raw_message
        }


        # ----------------------------------------------------
        # Send the email
        # ----------------------------------------------------
        response = gmail.users().messages().send(
            userId="me",
            body=request_body
        ).execute()


        # ----------------------------------------------------
        # Return Gmail's response
        # ----------------------------------------------------
        # The response contains information such as the
        # message ID created by Gmail.
        return response


    except ValueError:

        # ----------------------------------------------------
        # Re-raise validation errors
        # ----------------------------------------------------
        # These errors are already clear, so we don't change
        # their messages.
        raise


    except Exception as e:

        # ----------------------------------------------------
        # Handle unexpected Gmail API errors
        # ----------------------------------------------------
        # Add context so the caller knows that the failure
        # happened while sending the email.
        raise RuntimeError(
            f"Failed to send email to '{recipient}': {e}"
        ) from e