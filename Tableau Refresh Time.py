import requests
import xml.etree.ElementTree as ET
import pandas as pd
import argparse

def sign_in(server_url, pat_name, pat_secret, site_content_url=''):
    signin_url = f"{server_url}/api/3.4/auth/signin"
    payload = f"""
    <tsRequest>
      <credentials personalAccessTokenName="{pat_name}" personalAccessTokenSecret="{pat_secret}">
        <site contentUrl="{site_content_url}" />
      </credentials>
    </tsRequest>
    """
    headers = {'Content-Type': 'application/xml'}
    response = requests.post(signin_url, data=payload, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Sign-in failed: {response.text}")
    
    # Parse response, ignoring namespace
    response_text = response.text.replace('xmlns="http://tableau.com/api"', '')
    root = ET.fromstring(response_text)
    auth_token = root.find('.//credentials').get('token')
    site_id = root.find('.//site').get('id')
    return auth_token, site_id

def get_all_workbooks(server_url, site_id, auth_token):
    data = []
    page_number = 1
    page_size = 1000 # Max page size
    workbooks_url = f"{server_url}/api/3.4/sites/{site_id}/workbooks"
    
    while True:
        paged_url = f"{workbooks_url}?pageSize={page_size}&pageNumber={page_number}"
        headers = {'X-Tableau-Auth': auth_token}
        response = requests.get(paged_url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to get workbooks: {response.text}")
        
        # Parse response, ignoring namespace
        response_text = response.text.replace('xmlns="http://tableau.com/api"', '')
        root = ET.fromstring(response_text)
        pagination = root.find('pagination')
        total_available = int(pagination.get('totalAvailable'))
        
        for workbook in root.findall('.//workbook'):
            project = workbook.find('project')
            project_name = project.get('name') if project is not None else 'None'
            workbook_name = workbook.get('name')
            # Exclude workbooks and projects containing 'Archive'
            if 'Archive' not in project_name and 'Archive' not in workbook_name:
                updated_at = workbook.get('updatedAt')
                data.append({
                    'Project': project_name,
                    'Workbook': workbook_name,
                    'UpdatedAt': updated_at
                })
        
        if page_number * page_size >= total_available:
            break
        page_number += 1
    
    return data

def main():
    parser = argparse.ArgumentParser(description="Extract Tableau workbooks with projects and updated dates, excluding 'Archive'")
    parser.add_argument('--server_url', required=True, help='Tableau Server URL, e.g., https://your-tableau-server')
    parser.add_argument('--pat_name', required=True, help='Personal Access Token Name')
    parser.add_argument('--pat_secret', required=True, help='Personal Access Token Secret')
    parser.add_argument('--site_content_url', default='', help='Site content URL (leave empty for default site)')
    parser.add_argument('--output_file', default='tableau_workbooks.xlsx', help='Output Excel file name')
    
    args = parser.parse_args()
    
    auth_token, site_id = sign_in(args.server_url, args.pat_name, args.pat_secret, args.site_content_url)
    data = get_all_workbooks(args.server_url, site_id, auth_token)
    df = pd.DataFrame(data)
    df.to_excel(args.output_file, index=False)
    print(f"Data saved to {args.output_file}")

if __name__ == "__main__":
    main()
