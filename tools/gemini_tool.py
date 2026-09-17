# ============================================================
# Gemini Tool
# ============================================================
# Purpose:
#     This file contains all Gemini-related functionality
#     for our AI Email Agent.
#
# Current responsibility:
#     1. Initialize the Gemini client
#     2. Send a prompt to Gemini
#     3. Return Gemini's generated response
#
# Why separate Gemini code into a tool?
#     The agent should not need to know the details of how
#     the Gemini API works.
#
#     The agent can simply say:
#
#         generate_email(prompt)
#
#     and this tool handles the API communication.
#
# Error handling:
#     API errors are caught and given useful context.
# ============================================================

from google import genai


# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------

GEMINI_MODEL = "gemini-3.6-flash"


# ------------------------------------------------------------
# Create Gemini client
# ------------------------------------------------------------

def create_gemini_client(api_key):
    """
    Create and return a Gemini client.

    Parameters:
        api_key (str):
            Gemini API key.

    Returns:
        Gemini client object.
    """

    try:

        if not api_key or not str(api_key).strip():
            raise ValueError(
                "Gemini API key cannot be empty."
            )

        client = genai.Client(
            api_key=str(api_key).strip()
        )

        return client

    except ValueError:
        raise

    except Exception as e:
        raise RuntimeError(
            f"Failed to create Gemini client: {e}"
        ) from e


# ------------------------------------------------------------
# Generate text using Gemini
# ------------------------------------------------------------

def generate_text(client, prompt):
    """
    Send a prompt to Gemini and return the generated text.

    Parameters:
        client:
            Gemini client created by create_gemini_client().

        prompt (str):
            Instructions that will be sent to Gemini.

    Returns:
        Generated text as a string.
    """

    try:

        if not prompt or not str(prompt).strip():
            raise ValueError(
                "Gemini prompt cannot be empty."
            )

        response = client.interactions.create(
            model=GEMINI_MODEL,
            input=str(prompt)
        )

        if not response.output_text:
            raise RuntimeError(
                "Gemini returned an empty response."
            )

        return response.output_text

    except ValueError:
        raise

    except Exception as e:
        raise RuntimeError(
            f"Gemini API request failed: {e}"
        ) from e

# ------------------------------------------------------------
# Decide what action should be taken
# ------------------------------------------------------------

def decide_action(client, task, status, email_status):
    """
    Ask Gemini to decide what action should be taken
    for a task.

    Gemini must return exactly one of:

        send_reminder
        skip

    The Python application will validate the result before
    performing any real-world action.
    """

    try:

        # ----------------------------------------------------
        # Validate inputs
        # ----------------------------------------------------

        if not task or not str(task).strip():
            raise ValueError(
                "Task cannot be empty."
            )

        if not status or not str(status).strip():
            raise ValueError(
                "Task status cannot be empty."
            )


        # ----------------------------------------------------
        # Create the decision prompt
        # ----------------------------------------------------

        prompt = f"""
You are deciding what action an email automation agent
should take.

Task: {task}
Task status: {status}
Email status: {email_status}

Rules:

1. If the task status is not Pending, return:
   skip

2. If the email status is Sent, return:
   skip

3. Otherwise return:
   send_reminder

Return ONLY one of these exact values:

send_reminder
skip

Do not provide an explanation.
"""


        # ----------------------------------------------------
        # Ask Gemini for the decision
        # ----------------------------------------------------

        response = client.interactions.create(
            model=GEMINI_MODEL,
            input=prompt
        )


        # ----------------------------------------------------
        # Extract and clean the response
        # ----------------------------------------------------

        decision = response.output_text.strip().lower()


        # ----------------------------------------------------
        # Validate Gemini's decision
        # ----------------------------------------------------

        allowed_decisions = {
            "send_reminder",
            "skip"
        }

        if decision not in allowed_decisions:

            raise ValueError(
                f"Gemini returned an invalid decision: "
                f"{decision}"
            )


        return decision


    except ValueError:
        raise

    except Exception as e:

        raise RuntimeError(
            f"Gemini decision request failed: {e}"
        ) from e