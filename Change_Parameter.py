import os
import time
import datetime
import requests
import json
import pandas as pd  # Import pandas
from msal import ConfidentialClientApplication

# ----- Configurations (Replace with your actual values) -----
app_id = "cc1930f0-1a02-4691-97c5-776ebfca744d"  # Your Azure AD App Registration client ID
tenant_id = "e5f2178a-ceaa-4d58-9c63-906f15e01565"
app_secret = "Qbk8Q~eT7OzgItTBUi2AhdCZ_wYVZEVdWevigbA~"

excel_file = r"C:\Users\Exill\Downloads\Truespot 1.xlsx"
group_id = "2a93b9b2-eddd-46d2-993d-cb0e2df89ca1"
# -----------------------------------------------------------

# Service Principal Authentication
app = ConfidentialClientApplication(
    client_id=app_id,
    client_credential=app_secret,
    authority=f"https://login.microsoftonline.com/{tenant_id}"
)
result = app.acquire_token_for_client(scopes=["https://analysis.windows.net/powerbi/api/.default"])
access_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyIsImtpZCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTVmMjE3OGEtY2VhYS00ZDU4LTljNjMtOTA2ZjE1ZTAxNTY1LyIsImlhdCI6MTczNzAwNTg0MCwibmJmIjoxNzM3MDA1ODQwLCJleHAiOjE3MzcwMTAwNTYsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVlFBcS84WkFBQUFWVlVnYytWNllVTUxJdm1jMHFMcGFDeVVMNk5hYkRGU0Z3NTdoend1WDNOOERFTW1QU2JPaHlkSWw2a052b2J4M09lZ1BrVTJDd3JnNGZoU0RMN2FoZjY3czFuTkp5Z1cwTFlKUExrN1RnMD0iLCJhbXIiOlsicHdkIiwicnNhIiwibWZhIl0sImFwcGlkIjoiMThmYmNhMTYtMjIyNC00NWY2LTg1YjAtZjdiZjJiMzliM2YzIiwiYXBwaWRhY3IiOiIwIiwiZGV2aWNlaWQiOiJmODkyODQ3Yy03NTQ1LTRkNDYtYTFlYS02OGZkNWVkYjk3MDMiLCJmYW1pbHlfbmFtZSI6Ikd1cHRhIiwiZ2l2ZW5fbmFtZSI6IkphbnZpIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTIyLjE3MC4xMTIuOTgiLCJuYW1lIjoiSmFudmkgR3VwdGEiLCJvaWQiOiI2YzkwODMwNy1hZTcyLTQ2MzAtYWJkZS0xMmEzMzc4MmQ5OGEiLCJwdWlkIjoiMTAwMzIwMDM0QTRBNkZBOCIsInJoIjoiMS5BVW9BaWhmeTVhck9XRTJjWTVCdkZlQVZaUWtBQUFBQUFBQUF3QUFBQUFBQUFBQ0pBRGxLQUEuIiwic2NwIjoiQXBwLlJlYWQuQWxsIENhcGFjaXR5LlJlYWQuQWxsIENhcGFjaXR5LlJlYWRXcml0ZS5BbGwgQ29ubmVjdGlvbi5SZWFkLkFsbCBDb25uZWN0aW9uLlJlYWRXcml0ZS5BbGwgQ29udGVudC5DcmVhdGUgRGFzaGJvYXJkLlJlYWQuQWxsIERhc2hib2FyZC5SZWFkV3JpdGUuQWxsIERhdGFmbG93LlJlYWQuQWxsIERhdGFmbG93LlJlYWRXcml0ZS5BbGwgRGF0YXNldC5SZWFkLkFsbCBEYXRhc2V0LlJlYWRXcml0ZS5BbGwgR2F0ZXdheS5SZWFkLkFsbCBHYXRld2F5LlJlYWRXcml0ZS5BbGwgSXRlbS5FeGVjdXRlLkFsbCBJdGVtLkV4dGVybmFsRGF0YVNoYXJlLkFsbCBJdGVtLlJlYWRXcml0ZS5BbGwgSXRlbS5SZXNoYXJlLkFsbCBPbmVMYWtlLlJlYWQuQWxsIE9uZUxha2UuUmVhZFdyaXRlLkFsbCBQaXBlbGluZS5EZXBsb3kgUGlwZWxpbmUuUmVhZC5BbGwgUGlwZWxpbmUuUmVhZFdyaXRlLkFsbCBSZXBvcnQuUmVhZFdyaXRlLkFsbCBSZXBydC5SZWFkLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkV3JpdGUuQWxsIFRlbmFudC5SZWFkLkFsbCBUZW5hbnQuUmVhZFdyaXRlLkFsbCBVc2VyU3RhdGUuUmVhZFdyaXRlLkFsbCBXb3Jrc3BhY2UuR2l0Q29tbWl0LkFsbCBXb3Jrc3BhY2UuR2l0VXBkYXRlLkFsbCBXb3Jrc3BhY2UuUmVhZC5BbGwgV29ya3NwYWNlLlJlYWRXcml0ZS5BbGwi"
headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

last_modified_time = os.path.getmtime(excel_file)

# Adjusted interval checking
check_interval = 5
extended_check_interval = 15

while True:
    try:
        current_modified_time = os.path.getmtime(excel_file)
        if current_modified_time > last_modified_time:
            print("Excel file modified. Running update script...")
            last_modified_time = current_modified_time
        else:
            print("No changes detected.")
        
        # Switch intervals after the first check
        time.sleep(check_interval)
        check_interval = extended_check_interval

    except KeyboardInterrupt:
        print("Script stopped by the user.")
        break
    except Exception as e:
        print(f"Error: {e}")
