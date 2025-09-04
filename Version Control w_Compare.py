import urllib.request
import urllib.error
import json
from datetime import datetime, date, timedelta, timezone
import os
import glob
import re
import logging
import zipfile
import shutil
import ntpath
import xml.etree.ElementTree as ET
import sys
from collections import defaultdict
from typing import Dict, List, Set, Tuple, Any, Optional

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration - replace with your details
SERVER_URL = 'https://example.tableau.com' # e.g., https://example.tableau.com
API_VERSION = '3.4' # Current Tableau API version
TOKEN_NAME = '' # Name of your Personal Access Token
TOKEN_SECRET = '' # Personal Access Token
SITE_CONTENT_URL = '' # Empty string for default site, or the content URL like 'Marketing'
ALLOWED_PROJECTS = {'Test', 'My Reports'}

# Update SAVE_DIR to your network path or a local path for testing
SAVE_DIR = os.path.expanduser('~\Documents')

class TableauWorkbookComparator:
    def __init__(self):
        self.changes = defaultdict(list)
        
    def parse_workbook(self, file_path: str) -> ET.Element:
        """Parse a Tableau workbook XML file."""
        try:
            tree = ET.parse(file_path)
            return tree.getroot()
        except ET.ParseError as e:
            logger.error(f"Error parsing {file_path}: {e}")
            raise
        except FileNotFoundError:
            logger.error(f"File not found: {file_path}")
            raise
    
    def extract_worksheets(self, root: ET.Element) -> Dict[str, ET.Element]:
        """Extract all worksheets from the workbook."""
        worksheets = {}
        for worksheet in root.findall('.//worksheet'):
            name = worksheet.get('name', 'Unnamed')
            worksheets[name] = worksheet
        return worksheets
    
    def extract_dashboards(self, root: ET.Element) -> Dict[str, ET.Element]:
        """Extract all dashboards from the workbook."""
        dashboards = {}
        for dashboard in root.findall('.//dashboard'):
            name = dashboard.get('name', 'Unnamed')
            dashboards[name] = dashboard
        return dashboards
    
    def extract_datasources(self, root: ET.Element) -> Dict[str, ET.Element]:
        """Extract all datasources from the workbook."""
        datasources = {}
        datasources_container = root.find('datasources')
        if datasources_container is not None:
            for datasource in datasources_container.findall('datasource'):
                name = datasource.get('name', 'Unnamed')
                caption = datasource.get('caption', name)
                key = caption if caption and caption != name else name
                if key and key != 'Parameters':
                    datasources[key] = datasource
        for datasource in root.findall('.//datasource'):
            name = datasource.get('name', 'Unnamed')
            caption = datasource.get('caption', name)
            key = caption if caption and caption != name else name
            if key and key != 'Parameters' and key not in datasources:
                datasources[key] = datasource
        return datasources
    
    def extract_datasource_details(self, datasource: ET.Element) -> Dict[str, Any]:
        """Extract detailed information about a datasource."""
        details = {
            'connection_type': None,
            'server': None,
            'database': None,
            'tables': [],
            'columns': [],
            'relations': [],
            'custom_sql': None,
            'initial_sql': None,
            'connection_attributes': {}
        }
        connection = datasource.find('.//connection')
        if connection is not None:
            details['connection_type'] = connection.get('class')
            details['server'] = connection.get('server')
            details['database'] = connection.get('dbname') or connection.get('database')
            details['initial_sql'] = connection.get('initial-sql')
            details['connection_attributes'] = dict(connection.attrib)
            
            # Get custom SQL if present
            relation = connection.find('.//relation')
            if relation is not None and relation.get('type') == 'text':
                details['custom_sql'] = relation.text
        
        # Get tables/relations
        for relation in datasource.findall('.//relation'):
            if relation.get('table'):
                details['tables'].append({
                    'name': relation.get('table'),
                    'type': relation.get('type', 'table')
                })
        
        # Get columns
        for column in datasource.findall('.//column'):
            col_info = {
                'name': column.get('name'),
                'datatype': column.get('datatype'),
                'role': column.get('role'),
                'type': column.get('type')
            }
            details['columns'].append(col_info)
        
        return details
    
    def extract_parameters(self, root: ET.Element) -> Dict[str, ET.Element]:
        """Extract all parameters from the workbook."""
        parameters = {}
        param_datasource = root.find('.//datasource[@name="Parameters"]')
        if param_datasource is not None:
            for param in param_datasource.findall('.//column[@param="true"]'):
                name = param.get('name', 'Unnamed')
                parameters[name] = param
        return parameters
    
    def extract_calculated_fields(self, datasource: ET.Element) -> Dict[str, str]:
        """Extract calculated fields from a datasource."""
        calc_fields = {}
        for column in datasource.findall('.//column[@datatype]'):
            name = column.get('caption', '')
            calculation = column.find('calculation')
            if calculation is not None and calculation.get('class') == 'tableau':
                formula = calculation.get('formula', '')
                calc_fields[name] = formula
        return calc_fields
    
    def compare_sets(self, old_set: Set[str], new_set: Set[str], item_type: str):
        """Compare two sets and record additions/deletions."""
        added = new_set - old_set
        removed = old_set - new_set
        if added:
            self.changes[f'{item_type}_added'].extend(list(added))
        if removed:
            self.changes[f'{item_type}_removed'].extend(list(removed))
    
    def compare_element_attributes(self, old_elem: ET.Element, new_elem: ET.Element, 
                                  elem_name: str, item_type: str):
        """Compare attributes of XML elements."""
        old_attrs = old_elem.attrib
        new_attrs = new_elem.attrib
        for attr_name in set(old_attrs.keys()) | set(new_attrs.keys()):
            old_val = old_attrs.get(attr_name)
            new_val = new_attrs.get(attr_name)
            if old_val != new_val:
                change_desc = f"{elem_name}: {attr_name} changed from '{old_val}' to '{new_val}'"
                self.changes[f'{item_type}_modified'].append(change_desc)
    
    def compare_calculated_fields(self, old_fields: Dict[str, str], 
                                 new_fields: Dict[str, str], datasource_name: str):
        """Compare calculated fields between datasources."""
        old_field_names = set(old_fields.keys())
        new_field_names = set(new_fields.keys())
        added = new_field_names - old_field_names
        removed = old_field_names - new_field_names
        if added:
            for field in added:
                self.changes['calculated_fields_added'].append(f"{datasource_name}: {field}")
        if removed:
            for field in removed:
                self.changes['calculated_fields_removed'].append(f"{datasource_name}: {field}")
        common_fields = old_field_names & new_field_names
        for field in common_fields:
            if old_fields[field] != new_fields[field]:
                self.changes['calculated_fields_modified'].append(
                    f"{datasource_name}: {field} formula changed"
                )
    
    def compare_datasource_details(self, old_details: Dict[str, Any], 
                                  new_details: Dict[str, Any], datasource_name: str):
        """Compare detailed datasource information."""
        connection_fields = ['connection_type', 'server', 'database']
        for field in connection_fields:
            if old_details[field] != new_details[field]:
                self.changes['datasource_connections_modified'].append(
                    f"{datasource_name}: {field} changed from '{old_details[field]}' to '{new_details[field]}'"
                )
        if old_details['custom_sql'] != new_details['custom_sql']:
            if old_details['custom_sql'] and new_details['custom_sql']:
                self.changes['datasource_custom_sql_modified'].append(
                    f"{datasource_name}: Custom SQL modified"
                )
            elif new_details['custom_sql']:
                self.changes['datasource_custom_sql_added'].append(
                    f"{datasource_name}: Custom SQL added"
                )
            elif old_details['custom_sql']:
                self.changes['datasource_custom_sql_removed'].append(
                    f"{datasource_name}: Custom SQL removed"
                )
        if old_details['initial_sql'] != new_details['initial_sql']:
            if old_details['initial_sql'] and new_details['initial_sql']:
                self.changes['datasource_initial_sql_modified'].append(
                    f"{datasource_name}: Initial SQL modified"
                )
            elif new_details['initial_sql']:
                self.changes['datasource_initial_sql_added'].append(
                    f"{datasource_name}: Initial SQL added"
                )
            elif old_details['initial_sql']:
                self.changes['datasource_initial_sql_removed'].append(
                    f"{datasource_name}: Initial SQL removed"
                )
        old_tables = {t['name'] for t in old_details['tables']}
        new_tables = {t['name'] for t in new_details['tables']}
        added_tables = new_tables - old_tables
        removed_tables = old_tables - new_tables
        for table in added_tables:
            self.changes['datasource_tables_added'].append(f"{datasource_name}: {table}")
        for table in removed_tables:
            self.changes['datasource_tables_removed'].append(f"{datasource_name}: {table}")
        old_columns = {(c['name'], c['datatype']) for c in old_details['columns'] if c['name'] and c['datatype']}
        new_columns = {(c['name'], c['datatype']) for c in new_details['columns'] if c['name'] and c['datatype']}
        added_columns = new_columns - old_columns
        removed_columns = old_columns - new_columns
        for name, datatype in added_columns:
            self.changes['datasource_columns_added'].append(f"{datasource_name}: {name} ({datatype})")
        for name, datatype in removed_columns:
            self.changes['datasource_columns_removed'].append(f"{datasource_name}: {name} ({datatype})")
    
    def compare_workbooks(self, old_file: str, new_file: str) -> Dict[str, List[str]]:
        """Main comparison function."""
        logger.info(f"Comparing {old_file} with {new_file}")
        old_root = self.parse_workbook(old_file)
        new_root = self.parse_workbook(new_file)
        old_worksheets = self.extract_worksheets(old_root)
        new_worksheets = self.extract_worksheets(new_root)
        old_dashboards = self.extract_dashboards(old_root)
        new_dashboards = self.extract_dashboards(new_root)
        old_datasources = self.extract_datasources(old_root)
        new_datasources = self.extract_datasources(new_root)
        old_parameters = self.extract_parameters(old_root)
        new_parameters = self.extract_parameters(new_root)
        self.compare_sets(set(old_worksheets.keys()), set(new_worksheets.keys()), 'worksheets')
        common_worksheets = set(old_worksheets.keys()) & set(new_worksheets.keys())
        for ws_name in common_worksheets:
            old_ws = old_worksheets[ws_name]
            new_ws = new_worksheets[ws_name]
            self.compare_element_attributes(old_ws, new_ws, ws_name, 'worksheets')
        self.compare_sets(set(old_dashboards.keys()), set(new_dashboards.keys()), 'dashboards')
        common_dashboards = set(old_dashboards.keys()) & set(new_dashboards.keys())
        for db_name in common_dashboards:
            old_db = old_dashboards[db_name]
            new_db = new_dashboards[db_name]
            self.compare_element_attributes(old_db, new_db, db_name, 'dashboards')
        self.compare_sets(set(old_datasources.keys()), set(new_datasources.keys()), 'datasources')
        common_datasources = set(old_datasources.keys()) & set(new_datasources.keys())
        for ds_name in common_datasources:
            old_ds = old_datasources[ds_name]
            new_ds = new_datasources[ds_name]
            self.compare_element_attributes(old_ds, new_ds, ds_name, 'datasources')
            old_ds_details = self.extract_datasource_details(old_ds)
            new_ds_details = self.extract_datasource_details(new_ds)
            self.compare_datasource_details(old_ds_details, new_ds_details, ds_name)
            old_calc_fields = self.extract_calculated_fields(old_ds)
            new_calc_fields = self.extract_calculated_fields(new_ds)
            self.compare_calculated_fields(old_calc_fields, new_calc_fields, ds_name)
        self.compare_sets(set(old_parameters.keys()), set(new_parameters.keys()), 'parameters')
        common_parameters = set(old_parameters.keys()) & set(new_parameters.keys())
        for param_name in common_parameters:
            old_param = old_parameters[param_name]
            new_param = new_parameters[param_name]
            self.compare_element_attributes(old_param, new_param, param_name, 'parameters')
        return dict(self.changes)
    
    def print_summary(self, changes: Dict[str, List[str]]):
        """Print a formatted summary of changes."""
        print("\n" + "="*60)
        print("TABLEAU WORKBOOK COMPARISON SUMMARY")
        print("="*60)
        if not any(changes.values()):
            print("No changes detected between the workbooks.")
            return
        change_types = [
            ('worksheets_added', 'Worksheets Added'),
            ('worksheets_removed', 'Worksheets Removed'),
            ('worksheets_modified', 'Worksheets Modified'),
            ('dashboards_added', 'Dashboards Added'),
            ('dashboards_removed', 'Dashboards Removed'),
            ('dashboards_modified', 'Dashboards Modified'),
            ('datasources_added', 'Data Sources Added'),
            ('datasources_removed', 'Data Sources Removed'),
            ('datasources_modified', 'Data Sources Modified'),
            ('datasource_connections_modified', 'Data Source Connections Modified'),
            ('datasource_custom_sql_added', 'Custom SQL Added'),
            ('datasource_custom_sql_removed', 'Custom SQL Removed'),
            ('datasource_custom_sql_modified', 'Custom SQL Modified'),
            ('datasource_initial_sql_added', 'Initial SQL Added'),
            ('datasource_initial_sql_removed', 'Initial SQL Removed'),
            ('datasource_initial_sql_modified', 'Initial SQL Modified'),
            ('datasource_tables_added', 'Data Source Tables Added'),
            ('datasource_tables_removed', 'Data Source Tables Removed'),
            ('datasource_columns_added', 'Data Source Columns Added'),
            ('datasource_columns_removed', 'Data Source Columns Removed'),
            ('calculated_fields_added', 'Calculated Fields Added'),
            ('calculated_fields_removed', 'Calculated Fields Removed'),
            ('calculated_fields_modified', 'Calculated Fields Modified'),
            ('parameters_added', 'Parameters Added'),
            ('parameters_removed', 'Parameters Removed'),
            ('parameters_modified', 'Parameters Modified')
        ]
        for change_key, display_name in change_types:
            if change_key in changes and changes[change_key]:
                print(f"\n{display_name}:")
                for item in changes[change_key]:
                    print(f"  • {item}")
        total_changes = sum(len(items) for items in changes.values())
        print(f"\nTotal Changes: {total_changes}")
        print("="*60)

def sign_in(server, api_version, token_name, token_secret, site_content_url):
    try:
        url = f"{server}/api/{api_version}/auth/signin"
        payload = {
            "credentials": {
                "personalAccessTokenName": token_name,
                "personalAccessTokenSecret": token_secret,
                "site": {"contentUrl": site_content_url}
            }
        }
        req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), method='POST')
        req.add_header('Content-Type', 'application/json')
        req.add_header('Accept', 'application/json')
        logger.info(f"Signing in to {url}")
        with urllib.request.urlopen(req) as response:
            data = json.loads(response.read().decode('utf-8'))
            token = data['credentials']['token']
            site_id = data['credentials']['site']['id']
            logger.info("Successfully signed in")
            return token, site_id
    except urllib.error.HTTPError as e:
        logger.error(f"HTTP error during sign-in: {e.code} {e.reason}")
        raise
    except urllib.error.URLError as e:
        logger.error(f"URL error during sign-in: {e.reason}")
        raise
    except KeyError as e:
        logger.error(f"Key error in sign-in response: {e}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error during sign-in: {e}")
        raise

def get_all_workbooks(server, api_version, site_id, token):
    try:
        workbooks = []
        page_number = 1
        page_size = 100
        # Use UTC-aware datetime for comparison
        twenty_four_hours_ago = datetime.utcnow().replace(tzinfo=timezone.utc) - timedelta(hours=24)
        while True:
            url = f"{server}/api/{api_version}/sites/{site_id}/workbooks?pageSize={page_size}&pageNumber={page_number}"
            logger.info(f"Querying workbooks with URL: {url}")
            req = urllib.request.Request(url)
            req.add_header('Accept', 'application/json')
            req.add_header('X-Tableau-Auth', token)
            with urllib.request.urlopen(req) as response:
                data = json.loads(response.read().decode('utf-8'))
                logger.debug(f"API response: {json.dumps(data, indent=2)}")
                wb_list = data.get('workbooks', {}).get('workbook', []) if isinstance(data.get('workbooks'), dict) else data.get('workbooks', [])
                if not isinstance(wb_list, list):
                    logger.error(f"Expected 'workbooks' to be a list, got {type(wb_list)}")
                    return []
                for wb in wb_list:
                    if not isinstance(wb, dict):
                        logger.warning(f"Skipping invalid workbook entry: {wb}")
                        continue
                    project_name = wb.get('project', {}).get('name')
                    logger.debug(f"Workbook details: {wb}")
                    updated_at = wb.get('updatedAt')
                    if updated_at and project_name in ALLOWED_PROJECTS:
                        updated_dt = datetime.fromisoformat(updated_at.rstrip('Z') + '+00:00')
                        if updated_dt > twenty_four_hours_ago:
                            workbooks.append(wb)
                    else:
                        logger.info(f"Skipping workbook {wb.get('name', 'unknown')} in project {project_name} or no updatedAt")
                total = int(data.get('pagination', {}).get('totalAvailable', 0))
                logger.info(f"Fetched page {page_number}, {len(wb_list)} workbooks, {len(workbooks)} in allowed projects and recent, total: {total}")
                if not wb_list or len(wb_list) < page_size:
                    break
                page_number += 1
        return workbooks
    except urllib.error.HTTPError as e:
        try:
            error_body = e.read().decode('utf-8')
            logger.error(f"HTTP error fetching workbooks: {e.code} {e.reason}, details: {error_body}")
        except:
            logger.error(f"HTTP error fetching workbooks: {e.code} {e.reason}")
        raise
    except urllib.error.URLError as e:
        logger.error(f"URL error fetching workbooks: {e.reason}")
        raise
    except KeyError as e:
        logger.error(f"Key error in workbooks response: {e}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error fetching workbooks: {e}")
        raise

def add_long_path_prefix(path):
    """Add appropriate long path prefix for Windows paths."""
    if os.name != 'nt':
        return path
    if path.startswith('\\\\'):
        return f'\\\\?\\UNC{path[1:]}'
    elif not path.startswith('\\\\?\\'):
        return f'\\\\?\\{path}'
    return path

def download_workbook(server, api_version, site_id, token, wb_id, updated_at, wb_name, project_name):
    try:
        mod_date = datetime.fromisoformat(updated_at.rstrip('Z')).date().isoformat()
        wb_folder = os.path.join(SAVE_DIR, re.sub(r'[^\w\-]', '_', project_name), re.sub(r'[^\w\-]', '_', wb_name))
        wb_folder = add_long_path_prefix(wb_folder)
        logger.debug(f"Creating directory: {wb_folder}")
        os.makedirs(wb_folder, exist_ok=True)
        url = f"{server}/api/{api_version}/sites/{site_id}/workbooks/{wb_id}/content"
        logger.info(f"Downloading workbook {wb_name} from {url}")
        req = urllib.request.Request(url)
        req.add_header('X-Tableau-Auth', token)
        with urllib.request.urlopen(req) as response:
            headers = response.headers
            content_disp = headers['Content-Disposition']
            if not content_disp:
                raise ValueError("Content-Disposition header missing")
            filename_match = re.search(r'filename="([^"]+)"', content_disp)
            if not filename_match:
                raise ValueError("Could not parse filename from Content-Disposition")
            original_filename = filename_match.group(1)
            _, ext = os.path.splitext(original_filename)
            base = re.sub(r'[^\w\-]', '_', wb_name)
            new_filename = f"{base}_{mod_date}{ext}"
            save_path = os.path.join(wb_folder, new_filename)
            logger.debug(f"Saving workbook to: {save_path}")
            with open(save_path, 'wb') as f:
                f.write(response.read())
            logger.info(f"Downloaded workbook {wb_name} to {save_path}")
            return base, ext, new_filename, wb_folder
    except urllib.error.HTTPError as e:
        logger.error(f"HTTP error downloading workbook {wb_name}: {e.code} {e.reason}")
        raise
    except urllib.error.URLError as e:
        logger.error(f"URL error downloading workbook {wb_name}: {e.reason}")
        raise
    except OSError as e:
        logger.error(f"OS error downloading workbook {wb_name} to {save_path}: {e}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error downloading workbook {wb_name}: {e}")
        raise

def extract_twbx(twbx_path, wb_folder, base, mod_date):
    try:
        wb_folder = add_long_path_prefix(wb_folder) if not wb_folder.startswith('\\\\?\\') else wb_folder
        extract_path = os.path.join(wb_folder, "extracted")
        os.makedirs(extract_path, exist_ok=True)
        twb_files = []
        with zipfile.ZipFile(twbx_path, 'r') as zip_ref:
            logger.info(f"Inspecting contents of {twbx_path}")
            zip_contents = zip_ref.namelist()
            logger.debug(f"ZIP contents: {zip_contents}")
            for item in zip_contents:
                if item.endswith('.twb'):
                    sanitized_base = re.sub(r'[^\w\-]', '_', base)
                    dest_filename = f"{sanitized_base}_{mod_date}.twb"
                    dest_path = os.path.join(wb_folder, dest_filename)
                    logger.debug(f"Extracting .twb file to: {dest_path}")
                    with zip_ref.open(item) as source, open(dest_path, 'wb') as target:
                        shutil.copyfileobj(source, target)
                    twb_files.append(dest_path)
                    logger.info(f"Extracted and moved .twb file to: {dest_path}")
                else:
                    logger.debug(f"Skipping non-.twb file in archive: {item}")
        if os.path.exists(extract_path):
            shutil.rmtree(extract_path)
            logger.info(f"Deleted extracted folder: {extract_path}")
        if not twb_files:
            logger.warning(f"No .twb files found in {twbx_path}")
        return twb_files
    except zipfile.BadZipFile as e:
        logger.error(f"Failed to extract {twbx_path}: Invalid ZIP file: {e}")
        raise
    except OSError as e:
        logger.error(f"OS error extracting {twbx_path}: {e}")
        raise
    except Exception as e:
        logger.error(f"Error extracting {twbx_path}: {e}")
        raise

def manage_copies(base, ext, wb_folder):
    try:
        # Remove long path prefix for file operations
        wb_folder = wb_folder.replace('\\\\?\\UNC\\', '\\\\').replace('\\\\?\\', '') if wb_folder.startswith('\\\\?\\') else wb_folder
        logger.debug(f"Working directory: {wb_folder}")
        
        # Initialize extension for filtering
        ext_to_use = '.twb' if ext.lower() == '.twbx' else ext
        logger.debug(f"Using extension for filtering: {ext_to_use}")
        
        # Use glob to find all matching .twb files
        pattern = os.path.join(wb_folder, f"{base}_*.twb")
        logger.debug(f"Searching for files with pattern: {pattern}")
        all_files = glob.glob(pattern)
        logger.debug(f"All matching files found: {all_files}")
        
        file_dates = []
        for fn in all_files:
            bn = os.path.basename(fn)
            date_str = bn[len(base) + 1 : -len(ext_to_use)]
            logger.debug(f"Extracted date string from {bn}: {date_str}")
            try:
                if '_' in date_str:
                    date_str = date_str.replace('_', '-')
                d = datetime.fromisoformat(date_str).date()
                file_dates.append((d, fn))
            except ValueError as e:
                logger.warning(f"Skipping file with invalid date format {bn}: {date_str} (Error: {e})")
        file_dates.sort(key=lambda x: x[0], reverse=True)
        logger.info(f"Found {len(file_dates)} .twb files in {wb_folder}")
        
        # Keep up to 5 files, only if revision differs from the most recent
        to_delete = []
        if file_dates:
            latest_file = file_dates[0][1]
            latest_root = ET.parse(latest_file).getroot()
            latest_revision = latest_root.find('.//repository-location').get('revision')
            for date_, fn in file_dates[1:]:  # Skip the latest file
                root = ET.parse(fn).getroot()
                revision = root.find('.//repository-location').get('revision')
                if revision == latest_revision:
                    to_delete.append((date_, fn))
                else:
                    logger.info(f"Keeping {fn} due to different revision {revision} vs {latest_revision}")
        
        # Delete files with matching revisions beyond the top 5
        to_delete.sort(key=lambda x: x[0], reverse=True)
        for _, fn in to_delete[5:]:
            os.remove(fn)
            logger.info(f"Deleted old workbook copy with matching revision: {fn}")
        
        # Initialize changelog entry
        changelog_path = os.path.join(wb_folder, 'changelog.txt')
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        changelog_entry = f"\n\n=== Comparison on {current_date} ===\n"
        
        # Check if comparison is possible and revisions differ
        if len(file_dates) < 2:
            logger.info(f"Skipping comparison: Only {len(file_dates)} .twb file(s) found in {wb_folder}")
            return
        
        latest_file = file_dates[0][1]
        second_latest_file = file_dates[1][1]
        latest_root = ET.parse(latest_file).getroot()
        second_latest_root = ET.parse(second_latest_file).getroot()
        latest_revision = latest_root.find('.//repository-location').get('revision')
        second_latest_revision = second_latest_root.find('.//repository-location').get('revision')
        
        changes = {}  # Initialize changes dictionary
        if latest_revision != second_latest_revision:
            logger.info(f"Comparing workbooks due to revision change: {os.path.basename(second_latest_file)} (rev {second_latest_revision}) with {os.path.basename(latest_file)} (rev {latest_revision})")
            comparator = TableauWorkbookComparator()
            changes = comparator.compare_workbooks(second_latest_file, latest_file)
            
            # Update changelog only if changes are identified
            if any(changes.values()):
                # Prepare comparison output
                from io import StringIO
                output = StringIO()
                sys.stdout = output
                comparator.print_summary(changes)
                sys.stdout = sys.__stdout__
                comparison_text = output.getvalue()
                output.close()
                
                # Append comparison results to changelog
                changelog_entry += comparison_text
                existing_content = ""
                try:
                    with open(changelog_path, 'r') as f:
                        existing_content = f.read()
                except FileNotFoundError:
                    logger.info(f"Creating new changelog file: {changelog_path}")
                
                with open(changelog_path, 'w') as f:
                    f.write(changelog_entry + existing_content)
                logger.info(f"Changelog updated at: {changelog_path}")
        else:
            logger.info(f"Skipping comparison: No revision change between {os.path.basename(second_latest_file)} (rev {second_latest_revision}) and {os.path.basename(latest_file)} (rev {latest_revision})")
        
        # Handle file based on comparison result
        if not any(changes.values()) and latest_revision == second_latest_revision and os.path.exists(latest_file):
            os.remove(latest_file)
            logger.info(f"Deleted identical latest file with matching revision: {latest_file}")
        elif os.path.exists(latest_file):
            current_date = date.today().isoformat()
            new_filename = f"{os.path.splitext(latest_file)[0]}{ext_to_use}"
            os.rename(latest_file, new_filename)
            logger.info(f"Renamed different latest file to: {new_filename}")
    except Exception as e:
        logger.error(f"Error managing copies for {base}{ext} in {wb_folder}: {e}")
        raise

def main():
    try:
        logger.debug(f"Using SAVE_DIR: {SAVE_DIR}")
        if not os.path.exists(SAVE_DIR):
            logger.error(f"SAVE_DIR does not exist: {SAVE_DIR}")
            raise OSError(f"SAVE_DIR does not exist: {SAVE_DIR}")
        token, site_id = sign_in(SERVER_URL, API_VERSION, TOKEN_NAME, TOKEN_SECRET, SITE_CONTENT_URL)
        workbooks = get_all_workbooks(SERVER_URL, API_VERSION, site_id, token)
        if not workbooks:
            logger.info("No workbooks found in allowed projects")
            return
        for wb in workbooks:
            if not isinstance(wb, dict):
                logger.error(f"Invalid workbook entry: {wb}")
                continue
            wb_id = wb.get('id')
            name = wb.get('name')
            updated_at = wb.get('updatedAt')
            project_name = wb.get('project', {}).get('name')
            if not all([wb_id, name, updated_at, project_name]):
                logger.error(f"Missing required fields in workbook: {wb}")
                continue
            logger.info(f"Processing workbook {name} in project {project_name} modified at {updated_at}")
            base, ext, new_filename, wb_folder = download_workbook(SERVER_URL, API_VERSION, site_id, token, wb_id, updated_at, name, project_name)
            if ext.lower() == '.twbx':
                twb_files = extract_twbx(os.path.join(wb_folder, new_filename), wb_folder, base, mod_date=updated_at.rstrip('Z').split('T')[0])
                if twb_files:
                    manage_copies(base, '.twb', wb_folder)
                os.remove(os.path.join(wb_folder, new_filename))
                logger.info(f"Deleted original .twbx file: {new_filename}")
            else:
                manage_copies(base, ext, wb_folder)
        sign_out_url = f"{SERVER_URL}/api/{API_VERSION}/auth/signout"
        req = urllib.request.Request(sign_out_url, method='POST')
        req.add_header('X-Tableau-Auth', token)
        urllib.request.urlopen(req)
        logger.info("Signed out successfully")
    except Exception as e:
        logger.error(f"Error in main: {e}")
        raise

if __name__ == "__main__":
    main()