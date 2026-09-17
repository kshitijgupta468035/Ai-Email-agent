# ============================================================
# Excel Tool
# ============================================================
# Purpose:
#     This file contains all Excel-related functionality
#     for our AI Email Agent.
#
# Why separate Excel code into a tool?
#     Keeping Excel functionality separate makes our project
#     easier to maintain and allows the AI agent to use these
#     functions as tools later.
#
# Current responsibilities:
#     1. Read the Excel file
#     2. Validate required columns
#     3. Find pending tasks
#
# Error handling:
#     Errors are caught where we can add useful context.
#     We then raise the error so main.py / the agent can
#     decide what to do with it.
# ============================================================


# ------------------------------------------------------------
# Import Pandas
# ------------------------------------------------------------
# Pandas is used to read and process Excel files.
import pandas as pd


# ============================================================
# Function: read_tasks
# ============================================================
# Purpose:
#     Read an Excel file and return its contents as a
#     Pandas DataFrame.
#
# Parameters:
#     file_path (str):
#         Path to the Excel file.
#
# Returns:
#     pandas.DataFrame:
#         The data contained in the Excel file.
#
# Raises:
#     FileNotFoundError:
#         If the Excel file does not exist.
#
#     RuntimeError:
#         If another unexpected error occurs while reading
#         the Excel file.
#
# Example:
#     df = read_tasks("input/task.xlsx")
# ============================================================

def read_tasks(file_path):

    try:

        # ----------------------------------------------------
        # Read the Excel file
        # ----------------------------------------------------
        # pd.read_excel() loads the Excel data into a
        # Pandas DataFrame.
        df = pd.read_excel(file_path)

        # ----------------------------------------------------
        # Return the DataFrame
        # ----------------------------------------------------
        return df

    except FileNotFoundError:

        # ----------------------------------------------------
        # Handle missing Excel file
        # ----------------------------------------------------
        # Give a clear error message instead of exposing
        # a confusing low-level error.
        raise FileNotFoundError(
            f"Excel file not found: {file_path}"
        )

    except PermissionError:

        # ----------------------------------------------------
        # Handle permission problems
        # ----------------------------------------------------
        # This can happen if Python does not have permission
        # to access the file.
        raise PermissionError(
            f"Permission denied while accessing Excel file: "
            f"{file_path}"
        )

    except Exception as e:

        # ----------------------------------------------------
        # Handle unexpected errors
        # ----------------------------------------------------
        # We preserve the original error using "from e".
        # This is useful for debugging.
        raise RuntimeError(
            f"Failed to read Excel file '{file_path}': {e}"
        ) from e


# ============================================================
# Function: validate_task_columns
# ============================================================
# Purpose:
#     Make sure the Excel file contains all columns required
#     by our AI Email Agent.
#
# Required columns:
#     Name
#     Email
#     Task
#     Due_Date
#     Status
#
# Parameters:
#     df (pandas.DataFrame):
#         DataFrame that we want to validate.
#
# Returns:
#     True:
#         If all required columns are present.
#
# Raises:
#     ValueError:
#         If one or more required columns are missing.
#
# Example:
#     validate_task_columns(df)
# ============================================================

def validate_task_columns(df):

    # --------------------------------------------------------
    # Define the columns required by our agent
    # --------------------------------------------------------
    required_columns = [
        "Name",
        "Email",
        "Task",
        "Due_Date",
        "Status"
    ]

    # --------------------------------------------------------
    # Find missing columns
    # --------------------------------------------------------
    # We compare the required columns with the actual
    # columns present in the DataFrame.
    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    # --------------------------------------------------------
    # If any columns are missing, stop processing
    # --------------------------------------------------------
    if missing_columns:

        raise ValueError(
            f"Missing required columns: {missing_columns}"
        )

    # --------------------------------------------------------
    # All required columns are present
    # --------------------------------------------------------
    return True

# ============================================================
# Function: validate_task_data
# ============================================================
# Purpose:
#     Check whether the important fields in each task row
#     contain valid values before we send the task to Gemini
#     or Gmail.
#
# Required fields:
#     Name
#     Email
#     Task
#     Due_Date
#
# Parameters:
#     df (pandas.DataFrame):
#         DataFrame containing the task information.
#
# Raises:
#     ValueError:
#         If any required value is missing.
#
# Example:
#     validate_task_data(df)
# ============================================================

def validate_task_data(df):

    try:

        # ----------------------------------------------------
        # Columns that must contain a value
        # ----------------------------------------------------
        required_fields = [
            "Name",
            "Email",
            "Task",
            "Due_Date"
        ]

        # ----------------------------------------------------
        # Check every required field
        # ----------------------------------------------------
        for field in required_fields:

            # Find rows where the field is empty/missing.
            missing_rows = df[
                df[field].isna()
            ]

            # If we found any missing values, raise an error.
            if not missing_rows.empty:

                # Convert DataFrame indexes into Excel-like
                # row numbers.
                row_numbers = [
                    index + 2
                    for index in missing_rows.index
                ]

                raise ValueError(
                    f"Missing value in '{field}' "
                    f"at Excel row(s): {row_numbers}"
                )

        # ----------------------------------------------------
        # All required data is present
        # ----------------------------------------------------
        return True

    except ValueError:

        # Re-raise our validation error without changing it.
        raise

    except Exception as e:

        # Handle any unexpected validation error.
        raise RuntimeError(
            f"Failed to validate task data: {e}"
        ) from e

# ============================================================
# Function: get_pending_tasks
# ============================================================
# Purpose:
#     Return only tasks whose Status is "Pending".
#
# Parameters:
#     df (pandas.DataFrame):
#         DataFrame containing the task information.
#
# Returns:
#     pandas.DataFrame:
#         DataFrame containing only pending tasks.
#
# Example:
#     pending_tasks = get_pending_tasks(df)
# ============================================================

def get_pending_tasks(df):

    try:

        # ----------------------------------------------------
        # Make sure the Status column exists
        # ----------------------------------------------------
        # This protects the function if someone calls it
        # without validating the DataFrame first.
        if "Status" not in df.columns:

            raise ValueError(
                "Cannot find pending tasks because the "
                "'Status' column is missing."
            )

        # ----------------------------------------------------
        # Filter pending tasks
        # ----------------------------------------------------
        # This creates a True/False condition for every row.
        #
        # Status == "Pending" → True
        # Status == "Completed" → False
        #
        # Pandas keeps only the rows where the condition
        # is True.
        pending_tasks = df[
            df["Status"] == "Pending"
        ]

        # ----------------------------------------------------
        # Return pending tasks
        # ----------------------------------------------------
        return pending_tasks

    except ValueError:

        # ----------------------------------------------------
        # Re-raise validation errors
        # ----------------------------------------------------
        # We don't need to change the message because it is
        # already meaningful.
        raise

    except Exception as e:

        # ----------------------------------------------------
        # Handle unexpected errors
        # ----------------------------------------------------
        raise RuntimeError(
            f"Failed to find pending tasks: {e}"
        ) from e

def prepare_email_tracking_columns(df):
    """
    Make sure the Excel DataFrame contains the columns
    needed to track email processing.

    If the columns don't exist, they are created automatically.

    Returns:
        Updated DataFrame
    """

    try:

        if "Email_Status" not in df.columns:
            df["Email_Status"] = ""
        else: 
            # Column already exists. # Blank Excel cells may have been read as NaN. # Convert those NaN values to empty strings, # then make the column text-compatible. 
            df["Email_Status"] = ( 
                df["Email_Status"] 
                .fillna("not updated") 
                .astype(str) 
            )

        if "Last_Email_Sent" not in df.columns:
            df["Last_Email_Sent"] = ""
        else:
            # Column already exists. 
            # # Convert blank/missing values to empty strings 
            # # and make the column text-compatible. 
            df["Last_Email_Sent"] = ( 
                df["Last_Email_Sent"] 
                .fillna("not updated") 
                .astype(str) 
            )

        return df

    except Exception as e:

        raise RuntimeError(
            f"Failed to prepare email tracking columns: {e}"
        ) from e