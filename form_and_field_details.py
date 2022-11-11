import argparse
import json
from operator import itemgetter
import requests
import xlsxwritertools

url = 'https://z3n198.zendesk.com/api/v2/'

def build_request_session(email, token):
	uname = '{}/token'.format(email)
	pw = token
	headers = {'Content-Type': 'application/json'}
	auth = (uname, pw)
	session = requests.Session()
	session.auth = auth
	session.headers=headers
	return session

def get_field_info(session):
    endpoint = 'ticket_fields.json'
    full_url = url + endpoint
    ticket_fields = session.get(full_url)
    return ticket_fields.json()

def get_form_info(session):
    endpoint = 'ticket_forms.json'
    full_url = url + endpoint
    ticket_forms = session.get(full_url)
    return ticket_forms.json()

def load_json_data(fname):
    with open(fname, 'r') as f:
        jdata = json.load(f)
    return jdata

def get_fid_dict(ticket_fields):
    fid_dict = {fld['id']: fld for fld in ticket_fields['ticket_fields']}
    return fid_dict

def build_form_tab_data(ticket_forms, fid_dict):
    all_form_info = {}
    for form in ticket_forms['ticket_forms']:
        form_name = form['name']
        rows = []
        for fid in form['ticket_field_ids']:
            if fid in fid_dict:
                tix_field = fid_dict[fid]
                row = [
                        fid, 
                        tix_field['title'],
                        tix_field['type'], 
                        tix_field['required'],
                        tix_field['editable_in_portal'],
                ]
                if 'isInForm' not in tix_field:
                    tix_field['isInForm'] = []
                tix_field['isInForm'].append(form_name)
            else:
                row = [fid, 'Field does not exist', '', '', '']
            rows.append(row)
        all_form_info[form_name] = rows
    return all_form_info

def build_field_tab_data(fid_dict):
    fname_inform_dict = {}
    fname_notinform_dict = {}
    for fid, fdata in fid_dict.items():
        fname = fdata['title']
        if (fname in fname_inform_dict) or (fname in fname_notinform_dict):
            fname = '{} (2)'.format(fname)
        if 'isInForm' in fdata:
            fname_inform_dict[fname] = fdata
        else:
            fname_notinform_dict[fname] = fdata
    inform_field_rows = []
    for fname, fdata in sorted(fname_inform_dict.items(), key=itemgetter(0)):
            row = [
                    fdata['id'],
                    fname,
                    fdata['type'],
                    fdata['required'],
                    fdata['editable_in_portal'],
                    fdata['active'],
                    fdata['isInForm'],
            ]
            inform_field_rows.append(row)
    notinform_field_rows = []
    for fname, fdata in sorted(fname_notinform_dict.items(), key=itemgetter(0)):
            row = [
                    fdata['id'],
                    fname,
                    fdata['type'],
                    fdata['required'],
                    fdata['editable_in_portal'],
                    fdata['active'],
            ]
            notinform_field_rows.append(row)
    return inform_field_rows, notinform_field_rows

def write_spreadsheet(output_fname, inform_field_rows, notinform_field_rows, all_form_info):
    wb = xlsxwritertools.XLSXWorkbook(output_fname)
    #wb.build_default_styles()
    inf_sheet = wb.get_new_worksheet('Fields in Forms')
    col_dict = {0: {"label": "Field ID", "width": 10, "style": "idnum_style"},
            1: {"label": "Field Name", "width": 24, "style": "text_style"},
            2: {"label": "Type", "width": 14, "style": "text_style"},
            3: {"label": "Required", "width": 8, "style": "text_style"},
            4: {"label": "Portal Editable", "width": 15, "style": "text_style"},
            5: {"label": "Active", "width": 8, "style": "text_style"},
            6: {"label": "In Form", "width": 24, "style": "text_style", "multicolumn": True},
            }
    row = wb.fill_sheet(inf_sheet, col_dict, inform_field_rows)

    nif_sheet = wb.get_new_worksheet('Fields Not in Forms')
    col_dict = {0: {"label": "Field ID", "width": 10, "style": "idnum_style"},
            1: {"label": "Field Name", "width": 24, "style": "text_style"},
            2: {"label": "Type", "width": 14, "style": "text_style"},
            3: {"label": "Required", "width": 8, "style": "text_style"},
            4: {"label": "Portal Editable", "width": 15, "style": "text_style"},
            5: {"label": "Active", "width": 8, "style": "text_style"},
            }
    row = wb.fill_sheet(nif_sheet, col_dict, notinform_field_rows)

    form_col_dict = {0: {"label": "Field ID", "width": 10, "style": "idnum_style"},
            1: {"label": "Field Name", "width": 24, "style": "text_style"},
            2: {"label": "Type", "width": 14, "style": "text_style"},
            3: {"label": "Required", "width": 8, "style": "text_style"},
            4: {"label": "Portal Editable", "width": 15, "style": "text_style"},
            }
    for formname in sorted(all_form_info.keys()):
        sheet = wb.get_new_worksheet(formname)
        data = all_form_info[formname]
        row = wb.fill_sheet(sheet, form_col_dict, data)
    wb.close_workbook()

def parse_args():
    parser = argparse.ArgumentParser(
            description=('Get the field and ticket form info from a subdomain '
            'and put it into a spreadsheet'))
    parser.add_argument('--subdomain',
            type=str,
            help='Subdomain to query',
            required=False)
    parser.add_argument('--email',
            type=str,
            help='email to use for API call (not required)',
            required=False)
    parser.add_argument('--token', 
            type=str,
            help='API token to use for API calls (not required unless email is provided)',
            required=False)
    parser.add_argument('--field-file',
            type=str,
            help='path to JSON file holding field information',
            required=False,
            dest='field_file')
    parser.add_argument('--form-file',
            type=str,
            help='path to JSON file holding ticket form information',
            required=False,
            dest='form_file')
    parser.add_argument('--output',
            type=str,
            help="Name of the spreadsheet output file",
            required=True,
            dest="output")

    args = parser.parse_args()
    return args

def get_field_and_form_data(args):
    if args.email:
        if not args.token:
            e = "You must include a token along with an email to fetch data from the subdomain via API"
            raise Exception(e)
        session = build_request_session(args.email, args.token)
        ticket_fields = get_field_info(session)
        print('retrieved ticket fields')
        ticket_forms = get_form_info(session)
        print('retrieved ticket forms')
    elif args.field_file:
        if not args.form_file:
            e = "You must supply both a field file and a form file if you are not using the API"
            raise Exception(e)
        ticket_fields = load_json_data(args.field_file)
        ticket_forms = load_json_data(args.form_file)
    return ticket_fields, ticket_forms


if __name__ == "__main__":
    args = parse_args()
    url = 'https://{}.zendesk.com/api/v2/'.format(args.subdomain)
    print(url)
    ticket_fields, ticket_forms = get_field_and_form_data(args)
    fid_dict = get_fid_dict(ticket_fields)
    print('built fid_dict')
    all_form_info = build_form_tab_data(ticket_forms, fid_dict)
    print('built form tab info')
    inform_field_rows, notinform_field_rows = build_field_tab_data(fid_dict)
    print('built field tab info')
    write_spreadsheet(args.output, inform_field_rows, notinform_field_rows, all_form_info)
    print('Complete! Report written to {}'.format(args.output))
