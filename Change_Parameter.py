
# Approach 1:
import os
import time
import datetime
import requests
import json
import pandas as pd  # Import pandas
from msal import ConfidentialClientApplication

# ... (your other imports and code)
     # ----- Configurations (Replace with your actual values) -----
app_id = "cc1930f0-1a02-4691-97c5-776ebfca744d"  # Your Azure AD App Registration client ID
tenant_id = "e5f2178a-ceaa-4d58-9c63-906f15e01565"
app_secret = "Qbk8Q~eT7OzgItTBUi2AhdCZ_wYVZEVdWevigbA~"

excel_file = r"C:\Users\Exill\Downloads\Truespot 1.xlsx"
group_id = "2a93b9b2-eddd-46d2-993d-cb0e2df89ca1"
# -----------------------------------------------------------

# Service Principal Authentication (Recommended for unattended scripts)
# For interactive login, use PublicClientApplication with device code flow (see previous examples).
app = ConfidentialClientApplication(client_id=app_id, client_credential=app_secret, authority=f"https://login.microsoftonline.com/{tenant_id}" )
result = app.acquire_token_for_client(scopes=["https://analysis.windows.net/powerbi/api/.default"])
access_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyIsImtpZCI6InoxcnNZSEhKOS04bWdndDRIc1p1OEJLa0JQdyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTVmMjE3OGEtY2VhYS00ZDU4LTljNjMtOTA2ZjE1ZTAxNTY1LyIsImlhdCI6MTczNzAwNTg0MCwibmJmIjoxNzM3MDA1ODQwLCJleHAiOjE3MzcwMTAwNTYsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVlFBcS84WkFBQUFWVlVnYytWNllVTUxJdm1jMHFMcGFDeVVMNk5hYkRGU0Z3NTdoend1WDNOOERFTW1QU2JPaHlkSWw2a052b2J4M09lZ1BrVTJDd3JnNGZoU0RMN2FoZjY3czFuTkp5Z1cwTFlKUExrN1RnMD0iLCJhbXIiOlsicHdkIiwicnNhIiwibWZhIl0sImFwcGlkIjoiMThmYmNhMTYtMjIyNC00NWY2LTg1YjAtZjdiZjJiMzliM2YzIiwiYXBwaWRhY3IiOiIwIiwiZGV2aWNlaWQiOiJmODkyODQ3Yy03NTQ1LTRkNDYtYTFlYS02OGZkNWVkYjk3MDMiLCJmYW1pbHlfbmFtZSI6Ikd1cHRhIiwiZ2l2ZW5fbmFtZSI6IkphbnZpIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTIyLjE3MC4xMTIuOTgiLCJuYW1lIjoiSmFudmkgR3VwdGEiLCJvaWQiOiI2YzkwODMwNy1hZTcyLTQ2MzAtYWJkZS0xMmEzMzc4MmQ5OGEiLCJwdWlkIjoiMTAwMzIwMDM0QTRBNkZBOCIsInJoIjoiMS5BVW9BaWhmeTVhck9XRTJjWTVCdkZlQVZaUWtBQUFBQUFBQUF3QUFBQUFBQUFBQ0pBRGxLQUEuIiwic2NwIjoiQXBwLlJlYWQuQWxsIENhcGFjaXR5LlJlYWQuQWxsIENhcGFjaXR5LlJlYWRXcml0ZS5BbGwgQ29ubmVjdGlvbi5SZWFkLkFsbCBDb25uZWN0aW9uLlJlYWRXcml0ZS5BbGwgQ29udGVudC5DcmVhdGUgRGFzaGJvYXJkLlJlYWQuQWxsIERhc2hib2FyZC5SZWFkV3JpdGUuQWxsIERhdGFmbG93LlJlYWQuQWxsIERhdGFmbG93LlJlYWRXcml0ZS5BbGwgRGF0YXNldC5SZWFkLkFsbCBEYXRhc2V0LlJlYWRXcml0ZS5BbGwgR2F0ZXdheS5SZWFkLkFsbCBHYXRld2F5LlJlYWRXcml0ZS5BbGwgSXRlbS5FeGVjdXRlLkFsbCBJdGVtLkV4dGVybmFsRGF0YVNoYXJlLkFsbCBJdGVtLlJlYWRXcml0ZS5BbGwgSXRlbS5SZXNoYXJlLkFsbCBPbmVMYWtlLlJlYWQuQWxsIE9uZUxha2UuUmVhZFdyaXRlLkFsbCBQaXBlbGluZS5EZXBsb3kgUGlwZWxpbmUuUmVhZC5BbGwgUGlwZWxpbmUuUmVhZFdyaXRlLkFsbCBSZXBvcnQuUmVhZFdyaXRlLkFsbCBSZXBydC5SZWFkLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkLkFsbCBTdG9yYWdlQWNjb3VudC5SZWFkV3JpdGUuQWxsIFRlbmFudC5SZWFkLkFsbCBUZW5hbnQuUmVhZFdyaXRlLkFsbCBVc2VyU3RhdGUuUmVhZFdyaXRlLkFsbCBXb3Jrc3BhY2UuR2l0Q29tbWl0LkFsbCBXb3Jrc3BhY2UuR2l0VXBkYXRlLkFsbCBXb3Jrc3BhY2UuUmVhZC5BbGwgV29ya3NwYWNlLlJlYWRXcml0ZS5BbGwiLCJzaWduaW5fc3RhdGUiOlsia21zaSJdLCJzdWIiOiI3WThDMlRKcnZ5WEsySE1BQUpYWWZFck8tREdxWjE1Rmc1WjJONzZqM004IiwidGlkIjoiZTVmMjE3OGEtY2VhYS00ZDU4LTljNjMtOTA2ZjE1ZTAxNTY1IiwidW5pcXVlX25hbWUiOiJKYW52aUdAZXhpbGxhci5jb20iLCJ1cG4iOiJKYW52aUdAZXhpbGxhci5jb20iLCJ1dGkiOiJyVWJJczV1Mk8wS1ZTeFJ0NEw5dkFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiLCJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX2lkcmVsIjoiMTggMSJ9.OxA1nDx5NWDG3oft5CgqJ_nL-X4sZOrSawtDI2I_NnlGgMk0UEgBzqnTm10y4AMUdujEpbosE7DaKysz7LFvxO1aH5O97KzSI6k6KazmZ3Aury3-dr4c8wKnPdGVabE_YZVrg5x5g6wkFjb7QhMizzp3Dycza6cUxV2cVH1PZ_KS0OcjyORFq2xJ5Mnsm-lZdX-AqCssAZqxKeuckYbK1M05444VgR3DgHo5I9NHp6hmMyx4wHVycFvbBWdOiQkGdyrFbzEOMPZ8shpxzwTbobqZ_SdOEG1adgcYQD2luVMSBUt9ZyQosBfAfalFsi05ywLZ917iKQjZiW7OM3vuWg"
headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

last_modified_time = os.path.getmtime(excel_file)

while True:
    try:
        current_modified_time = os.path.getmtime(excel_file)
        if current_modified_time > last_modified_time:
            print("Excel file modified. Running update script...")
            last_modified_time = current_modified_time

            try:
                # Load the Excel workbook using pandas
                df = pd.read_excel(excel_file)  # Reads the entire Excel file into a DataFrame
                print("Reading Excel file..")
                for index, row in df.iterrows():  # Iterate through rows of the DataFrame
                    # Access column values by name (more readable and less error-prone)

                    #Get or create report dynamically
                    report_name = f"{row['Name']}" # Create report name dynamically 
                    # print(report_name)
                    reports = requests.get(f"https://api.powerbi.com/v1.0/myorg/groups/{group_id}/reports", headers=headers).json()['value']
                    # print(reports)
                    current_report = next((report for report in reports if report['name'] == report_name), None)
                    # print(current_report)
                    if current_report:
                        dataset_id = current_report['datasetId']
                        print("Dataset ID-> ",dataset_id)
                        parameters_payload = {
                            "updateDetails": [
                                {"name": "Minute Timezone offset with respect to UTC", "newValue": row["Minute Timezone offset with respect to UTC"]},
                                {"name": "Hospital Label", "newValue": row["Hospital Label"]},
                                {"name": "CustomerName", "newValue": row["CustomerName"]},
                                {"name": "Hour Timezone offset with respect to UTC", "newValue": row["Hour Timezone offset with respect to UTC"]},
                                {"name": "PostAggQueryPostfix", "newValue": row["PostAggQueryPostfix"]},
                                {"name": "AssteDIMQueryPostfix", "newValue": row["AssteDIMQueryPostfix"]},
                                {"name": "PreviousN", "newValue": row["PreviousN"]}
                            ]
                        }
            # fe7db397-5697-4582-af12-ab8e49e768de
                        api_url = f"https://api.powerbi.com/v1.0/myorg/groups/{group_id}/datasets/{dataset_id}/Default.UpdateParameters"
                        response = requests.post(api_url, json=parameters_payload, headers=headers)

                        if response.status_code == 200:
                            print(f"Parameters updated successfully for report '{report_name}'!")
                            print(f"Error: {response.status_code}")
                            print(response.headers)     # Print the headers
                            print(response.text)
                        else:
                            print(f"Failed to update parameters for report '{report_name}': {response.status_code}")
                            print(response.json())
                    else:
                        print(f"Report '{report_name}' not found in the workspace")



            except Exception as e:
                print(f"An error occurred: {e}")



        time.sleep(5)  # Check every 60 seconds (adjust as needed)

    except KeyboardInterrupt:  # Allow Ctrl+C to stop the script
        break    
    except Exception as e: # Handle any errors gracefully
        print(f"Error: {e}")




