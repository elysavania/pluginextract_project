import json
import pandas as pd
import pandasql as pdsql
import xlsxwritertools
from openpyxl import workbook
from decimal import Decimal
from autofit_spreadsheet_columns import autofit_spreadsheet_columns

# 2022-11-11 NOJ: Since Python 2.1 you can use \N{name} escape sequence to insert Unicode characters by their names.
# Source: https://stackoverflow.com/a/20799954

unicode_symbols = {
    "closing quotation": "\N{Right Double Quotation Mark}",
    "contain": "\N{Superset of or Equal To}",
    "cross": "\N{Cross Mark}",
    "does not contain": "\N{Not a Superset of}",
    "equal": "\N{Equals Sign}",
    "greater than or equal": "\N{Greater-Than or Slanted Equal To}",
    "less than or equal": "\N{Less-Than or Slanted Equal To}",
    "not equal": "\N{Not Equal To}",
    "opening quotation": "\N{Left Double Quotation Mark}",
    "tick": "\N{White Heavy Check Mark}",
}

crmrequirements = [
    "Salesforce Requirements to Connect the Plugin",
    "Salesforce must authorize Outreach through a Salesforce system user and meet the following requirements:",
    "The Salesforce system user must be able to modify data (create, edit, delete) on required objects that need to be shown in Outreach (i.e. Accounts, Contacts, Leads, Opportunities, User, User Role, Task/Event).",
    "The Salesforce system user must have Field Level Security settings that allow it to view and modify any mapped fields",
    "The profile the connecting Salesforce to Outreach has \"API Enabled\" under System Permissions in the Profile of the User",
    "Before configuring the bi-directional sync with Outreach, there are a few minimum requirements needed to leverage the Salesforce connection.",
    "The profile connecting Salesforce to Outreach can create or edit all objects (like Accounts, Contacts, Leads, Users, etc.)",
    "",
    "Outreach Requirements",
    "Outreach is compatible with Salesforce Lightning, Aloha (\"Classic\"), Console, and the SKUID overlay",
    "The Outreach user must be listed as an Admin within Outreach to have access to the plugin settings for connection.",
    "Outreach uses Rest API calls to communicate and sync with Salesforce. Enterprise & Unlimited editions of Salesforce are bundled with Rest API calls, but the Professional Edition is not.: https://developer.salesforce.com/docs/atlas.en-us.api_rest.meta/api_rest/",
    "If you are using Salesforce Professional Edition, you need to have the Web API Package and purchase API call bundles.",
    "To determine if your organization has purchased the API package, click on: Setup > Monitor > System Overview > API usage.",
    "To verify which version of Salesforce your company is using, follow these steps.",
    "If you are an existing Salesforce customer who is not on one of the above supported version and want to upgrade, contact your Salesforce Account Executive.",
    "",
    "SFDC Requirements: https://support.outreach.io/hc/en-us/articles/218582707"
]

# label_mapping includes all columns in the excel document
label_mapping = {
    "PluginID": "Plugin ID",
    "Provider": "Provider",
    "ProviderBaseURL": "{provider} Base URL",
    "OutreachBaseURL": "Outreach Base URL",
    "PluginAuthMode": "{provider} Auth Mode",
    "GlobalAPICallThreshold": "Global API call threshold",
    "OutreachSpecificAPICallThreshold": "Outreach-specific API call threshold",
    "PollUsersOnReconnect": "Refresh users on reconnect toggle",
    "InternalType": "Outreach Name",
    "ExternalType": "{provider} Name",
    "PollingEnabled": "POLLING: Periodically poll {provider} for new and changed",
    "PollingIntervalMinutes": "POLLING: Polling Frequency (min)",
    "PollingConditions": "POLLING: Polling Conditions",
    "MergeAndDeletePollingEnabled": "Merge & Delete Polling",
    "MergeAndDeletePollingFrequencyMinutes": "M & D Polling Freq",
    "SkipInboundDeleteInSequence": "Skip Inbound Delete in Sequence",
    "TrashBackfillEnabled": "Trash Backfill Enabled",
    "InboundCreateEnabled": "INBOUND CREATE: Create New {internal_type}s",
    "InboundCreateConditions": "INBOUND CREATE: Inbound Create Conditions",
    "InboundCreateContacts": "INBOUND CREATE: Create associated Contacts, Leads and Accounts",
    "InboundSyncAfterManualCreate": "INBOUND CREATE: Sync data down after manual create inside Outreach",
    "InboundUpdateEnabled": "INBOUND UPDATE: Update Existing {internal_type}s",
    "InboundUpdateConditions": "INBOUND UPDATE: Inbound Update Conditions",
    "OutboundCreateEnabled": "OUTBOUND CREATE: Create New {external_type}",
    "OutboundCreateConditions": "OUTBOUND CREATE: Outbound Create Conditions",
    "OutboundUpdateEnabled": "OUTBOUND UPDATE: Update Existing {external_type}",
    "OutboundUpdateConditions": "OUTBOUND UPDATE: Outbound Update Conditions",
    "OutboundPush": "PUSHING: Automatically push changes to CRM",
    "InternalField": "Outreach {internal_type} Field",
    "ExternalField": "{provider} {external_type} Field",
    "StandardTaskMappings": "MESSAGES & EVENTS",
    "OutboundCreateInMessages": "Inbound Messages",
    "OutboundCreateInMessagesCustom": "Inbound Messages: Customize Title",
    "OutboundCreateInMessagesTemplate": "Inbound Messages: Title Template",
    "OutboundCreateMailings": "Outbound Messages sent within Outreach",
    "OutboundCreateMailingsCustom": "Outbound Messages sent within Outreach: Customize Title",
    "OutboundCreateMailingsTemplate": "Outbound Messages sent within Outreach: Title Template",
    "OutboundCreateOutMessages": "Outbound Messages sent outside Outreach",
    "OutboundCreateOutMessagesCustom": "Outbound Messages sent outside Outreach: Customize Title",
    "OutboundCreateOutMessagesTemplate": "Outbound Messages sent outside Outreach: Title Template",
    "OutboundCreateNotes": "Notes",
    "OutboundCreateNotesCustom": "Notes: Customize Title ",
    "OutboundCreateNotesTemplate": "Notes: Title Template",
    "OutboundCreateCompletedTasks": "Completed Tasks",
    "OutboundCreateCompletedTasksCustom": "Completed Tasks: Customize Title",
    "OutboundCreateCompletedTasksTemplate": "Completed Tasks: Title Template",
    "OutboundCreateCalls": "Calls",
    "OutboundCreateCallsCustom": "Calls: Customize Title",
    "OutboundCreateCallsTemplate": "Calls: Title Template",
    "OutboundCreateCallsRecording": "Calls: Include call recordings in desscription",
    "OutboundCreateMailingClicks": "Message Clicks",
    "OutboundCreateMailingClicksCustom": "Message Clicks: Customize Title",
    "OutboundCreateMailingClicksTemplate": "Message Clicks: Title Template",
    "OutboundCreateMailingOpens": "Message Opens",
    "OutboundCreateMailingOpensCustom": "Message Opens: Customize Title",
    "OutboundCreateMailingOpensTemplate": "Message Opens: Title Template",
    "OutboundMessageIDField": "Include Message ID field for events: {provider} Field",
    "OutboundIncludeMessageID": "Include Message ID field for events"
}

# field_mapping indicates selected fields that are mapped in the CRM
field_mapping = {
    "InternalField": "Outreach Field Name",
    "InternalDefault": "Internal Default",
    "ExternalField": "SF Field Name",
    "ExternalMappedType": "External Mapped Type",
    "ExternalDefault": "External Default",
    "MappedField": "MappedField",
    "LookForNameInsteadOfID": "Look For Name Instead Of ID",
    "DisplayNameInsteadOfID": "Display Name Instead Of ID",
    "InboundEnabled": "Inbound Enabled",
    "OutboundOmitIfEmpty": "Outbound Omit If Empty",
    "InboundOmitIfEmpty": "Inbound Omit If Empty",
    "OutboundEnabled": "Outbound Enabled"
}

# preset_data_lead is the data captured from the lead config in the JSON file
preset_data_lead = {
    "account name": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Name of the account"
    },
    "actively being sequenced": {
        "FieldType": "Checkbox",
        "Outreach Engagement": "Record Data",
        "Note": "This field identifies if a Prospect is Active in a sequence."
    },
    "add date": {
        "FieldType": "Date/Time",
        "RecordType": "Record Data",
        "Note": "Date Prospect was added in Outreach? Not visible on the prospect page"
    },
    "first_name": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "First Name of Prospect"
    },
    "last_name": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Last name of Prospect"
    },
    "title": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Title of prospect"
    },
    "company": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Prospect's company"
    },
    "website": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Website URL"
    },
    "work_phone": {
        "FieldType": "Number",
        "RecordType": "Record Data",
        "Note": "Work Number"
    },
    "email": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Prospect's 1st email address"
    },
    "emails_opted_out": {
        "FieldType": "Checkbox",
        "RecordType": "Opt-Out",
        "Note": "Email opt out state confirmation (Only when granular opt out is enabled)"
    },
    "stage": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Prospect Stage in Outreach"
    },
    "owner": {
        "FieldType": "Lookup",
        "RecordType": "Record Data",
        "Note": "Owner of prospect in Outreach"
    },
    "address_state": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Prospect's state"
    },
    "source": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Prospect Source"
    },
    "address_street": {
        "FieldType": "Text",
        "RecordType": "Record Data",
        "Note": "Prospect's primary address"
    }
}
preset_data_contact = {}

preset_data_engagement_panel_field_note = "Untick both 'Skip empty values...' checkboxes in Advanced Field Mapping"
preset_data_engagement_panel_fields = dict.fromkeys([
    "current_sequence_name",
    "current_sequence_status",
    "current_sequence_step",
    "current_sequence_step_type",
    "current_sequence_id",
    "current_sequence_user_sfdc_id",
    "current_sequence_step_due",
    "current_sequence_id",
    "current_sequence_name",
    "current_sequence_user_sfdc_id",
    "current_sequence_status",
    "current_sequence_step",
    "current_sequence_step_type",
    "current_sequnce_step_due",
], {
    "FieldType": "",
    "RecordType": "",
    "Note": preset_data_engagement_panel_field_note
}
)

# dictionary cmd for leads
types_mapping_to_preset_data = {
    # 2022-11-14 NOJ: https://stackoverflow.com/questions/38987/how-do-i-merge-two-dictionaries-in-a-single-expression/26853961#26853961
    "Lead": dict(preset_data_lead, **preset_data_engagement_panel_fields),
    "Contact": dict(preset_data_contact, **preset_data_engagement_panel_fields),
}

# helper function to load plugin config json file


def read_plugin_json(fname="sage_plugin_configuration.json"):
    with open(fname, 'r') as f:
        plugin_data = json.load(f)
    return plugin_data

# To identify the plugin types and fields associated w/ the types


def get_mappings_dict(plugin_data):
    ptype_mappings = plugin_data['Legacy'].get('PluginTypeMappings', [])
    types = {}
    type_names = []
    for ptype in ptype_mappings:
        for temp in ptype['FieldMappings']:
            for key in temp:
                if key == 'Template':
                    temp['InternalField'] = temp[key]

        if ptype['InternalType'] == 'MessengerGroup':  # Ignore MessengerGroup object
            continue
        # name = str(ptype['ExternalType'])+'-'+str(ptype['InternalType'])
        name = (ptype['ExternalType'], ptype['InternalType'])
        type_names.append(name)
        types[name] = {"output": {}, "input": ptype}
    limits = plugin_data['Legacy']
    del limits['PluginTypeMappings']
    return limits, type_names, types

# To replace the provider with Provider


def update_provider_in_label_mapping(datadict):
    provider = datadict['Provider'].capitalize()
    for i in label_mapping:
        label_mapping[i] = label_mapping[i].replace(
            "{provider}", provider)

# To add types to label mappings


def update_external_internal_in_label_mapping(datalist, lm):
    external_type = datalist[0]
    internal_type = datalist[1]
    for index in lm:
        lm[index] = lm[index].replace("{external_type}", external_type)
        lm[index] = lm[index].replace("{internal_type}", internal_type)
    return (lm)

# To remove the label mappings and keep the mapping data


def update_labels_in_dictdata(data, lm):
    for label in lm:
        if label in data:
            data[lm[label]] = data.pop(label)
    return (data)

# To return the label mapping with its value


def update_label(value, lm):
    if value in lm:
        return lm[value]
    else:
        return (value)

# To search for the labels -> return values of the labels


def update_labels_in_list(lst, lm):
    for i in range(len(lst)):
        if lst[i] in lm:  # looks for list[index] in lm
            lst[i] = lm[lst[i]]
    return (lst)

# To intersperse an item in a list


def intersperse(lst, item):
    result = [item] * (len(lst) * 2 - 1)
    result[0::2] = lst
    return result


def surround_with_quotation_marks(value):
    return f"{unicode_symbols['opening quotation']}{value}{unicode_symbols['closing quotation']}"

# To create condition columns and add the mapping values to the columns


def write_conditions(value, label_mapping, row):
    logical_operator = (update_label(
        value['LogicalOperator'].upper(), label_mapping), '', '')
    if 'Conditions' in value.keys():
        conditions = value['Conditions']
    elif 'ConditionGroups' in value.keys():
        # Combined components of conditions into a single cell, replaced operator labels with symbols, added quotation marks to comparison values, and tidied up.
        # print(value["ConditionGroups"][0])
        row = write_conditions(value["ConditionGroups"][0], label_mapping, row)
        return row
    column_names = list(conditions[0].keys())
    order = [1, 0, 2]
    column_names = [column_names[i] for i in order]
    column_names[1] = update_label(column_names[1], unicode_symbols)

    df = pd.DataFrame(columns=column_names)
    df = df.append(conditions, ignore_index=True)
    df.fillna('null', inplace=True)
    list_of_conditions = df.values.tolist()
    # print(list_of_conditions)
    list_of_conditions = intersperse(list_of_conditions, logical_operator)
    # print(list_of_conditions)

    # rewrite the below - no for loop
    for i in list_of_conditions:
        if i[1] == '':
            # row = wb.add_single_row_shift(
            #     sheet, row, col_dict_conditions_operator, 1, update_labels_in_list(i, label_mapping))
            row = wb.add_single_row(
                sheet, row, col_dict_conditions, ("", f"{i[0]}"))
        else:
            i = [x if x != '' else 'null' for x in i]  # replace - with null
            # row = wb.add_single_row_shift(
            #     sheet, row, col_dict_conditions, 1, update_labels_in_list(i, label_mapping))
            field_name = i[0]
            field_name_with_quotes = surround_with_quotation_marks(field_name)
            operator_label = i[1]
            operator_symbol = update_label(operator_label, unicode_symbols)
            comparison_value = i[2]
            comparison_value_with_quotes = surround_with_quotation_marks(
                comparison_value)
            row = wb.add_single_row(
                sheet, row, col_dict_conditions, ("", f"{field_name_with_quotes} {operator_symbol} {comparison_value_with_quotes}"))
    if 'ConditionGroups' in value:
        # row = wb.add_single_row_shift(
        #     sheet, row, col_dict_conditions_operator, 1, logical_operator)
        row = wb.add_single_row(
            sheet, row, col_dict_conditions, ("", logical_operator[0]))
        row = write_conditions(value["ConditionGroups"][0], label_mapping, row)
    row = row + 1  # add a row after the condition)
    return row


# Below are all styling the sheet
if __name__ == "__main__":
    plain_text = {
        'width': 200,
        'style': 'text_style'
    }
    header_text = {
        'width': 200,
        'style': 'hdr_style',
        'text_wrap': True
    }
    url_text = {
        'width': 200,
        'style': 'url_style'
    }

    col_dict_level_0 = {
        # Field column style
        0: {
            'label': 'Field',
            'width': 50,
            'style': 'bold_style'
        },
        1: {
            'label': 'Value',
            'width': 50,
            'style': 'text_style'
        },

    }

    col_condition = {

        0: {
            'width': 200,
            'height': 100,
            'y_offset': 10,
            'x_offset': 10,
            'border': 1
        }
    }

    # Style for Messages & Events settings
    col_dict_task_mapping = {
        0: {
            'label': 'Field',
            'width': 100,
            'style': 'bold_text_style'
        },
        1: {
            'label': 'Value',
            'width': 100,
            'style': 'text_style'
        },

    }
    # Style for conditional operators
    col_dict_conditions_operator = {
        0: {
            'width': 50,
            'style': 'color_bold_text_style'
        },
        1: {'width': 50,
            'style': 'color_bold_text_style'
            },
        2: {'width': 50,
            'style': 'color_bold_text_style'
            }

    }
    # Style for conditions settings
    col_dict_conditions = {
        0: {
            'label': 'Field',
            'width': 50,
            'style': 'text_style'
        },
        1: {
            'label': 'Comparison Operator',
            'width': 50,
            'style': 'color_bold_text_style',
            'note': "Does Not Contain"
        }
    }
    col_field_mapping = {
        0: {
            'label': 'Internal Field',
            'width': 30,
            'style': 'text_style'
        },
        1: {
            'label': 'Internal Empty Placeholder',
            'width': 30,
            'style': 'text_style'
        },
        2: {
            'label': 'External Field',
            'width': 30,
            'style': 'text_style'
        },
        3: {
            'label': 'External Mapped Type',
            'width': 30,
            'style': 'text_style'
        },
        4: {
            'label': 'External Empty Placeholder',
            'width': 30,
            'style': 'text_style'
        },
        5: {
            'label': 'Mapped Field',
            'width': 30,
            'style': 'text_style'
        },
        6: {
            'label': 'Look For Name Instead Of record ID',
            'width': 30,
            'style': 'text_style'
        },
        7: {
            'label': 'Display Name Instead Of record ID',
            'width': 30,
            'style': 'text_style'
        },
        8: {
            'label': 'Updates In',
            'width': 30,
            'style': 'text_style'
        },
        9: {
            'label': 'Outbound Omit If Empty',
            'width': 30,
            'style': 'text_style'
        },
        10: {
            'label': 'Inbound Omit If Empty',
            'width': 30,
            'style': 'text_style'
        },
        11: {
            'label': 'Updates Out',
            'width': 30,
            'style': 'text_style'
        },
    }
    col_field_mapping1 = {
        0: {
            'label': 'Outreach Field Name',
            'width': 30,
            'style': 'text_style'
        },
        1: {
            'label': 'SF Field Name',
            'width': 30,
            'style': 'text_style'
        },
        2: {
            'label': 'Outreach Field Type',
            'width': 30,
            'style': 'text_style',
            'dropdown': [
                'Text',
                'Number',
                'Checkbox',
                'Date/Time',
                'Text (/Picklist)',
                'Lookup'
            ]
        },
        3: {
            'label': 'Outreach Record Type',
            'width': 30,
            'style': 'text_style',
            'dropdown': [
                'Record Data',
                'Opt-Out',
                'Outreach Engagement',
                'Custom Fields'
            ]
        },

        4: {
            'label': 'Internal Empty Placeholder',
            'width': 30,
            'style': 'text_style'
        },

        5: {
            'label': 'External Mapped Type',
            'width': 30,
            'style': 'text_style'
        },

        6: {
            'label': 'External Empty Placeholder',
            'width': 30,
            'style': 'text_style'
        },

        7: {
            'label': 'Mapped Field',
            'width': 30,
            'style': 'text_style'
        },

        8: {
            'label': 'Look For Name Instead Of record ID',
            'width': 30,
            'style': 'text_style'
        },

        9: {
            'label': 'Display Name Instead Of record ID',
            'width': 30,
            'style': 'text_style'
        },

        10: {
            'label': 'Updates In (SFDC > OR)',
            'width': 35,
            'style': 'color_checkboxes',
            'note': 'Updates In = Sync data from Salesforce to Outreach. When the box is unchecked, the field can be synced from Salesforce. When the box is checked, the field is selected to be synced from Salesforce. When there is no checkbox, the field only syncs to Salesforce.'
        },

        11: {
            'label': 'Updates Out (OR > SFDC)',
            'width': 35,
            'style': 'color_checkboxes',
            'note': 'Updates Out = Push data from Outreach to Salesforce. When the box is unchecked, the field can be synced to Salesforce. When the box is checked, the field is selected to be synced to Salesforce. When there is no checkbox, the field only syncs from Salesforce.'
        },

        12: {
            'label': 'Notes',
            'width': 30,
            'style': 'text_style'
        }
    }

    plugin_data = read_plugin_json()
    limits, type_names, types = get_mappings_dict(plugin_data)
    update_provider_in_label_mapping(limits)
    # Create the workbook

    spreadsheet_filename = 'sage_plugin_configuration.xlsx'
    wb = xlsxwritertools.XLSXWorkbook(spreadsheet_filename)

    # Create CRM Requirements Sheet
    sheet = wb.get_new_worksheet("CRM Requirements")
    i = 0
    for line in crmrequirements:
        if i in (0, 8, 17):
            wb.add_single_row_new_way(sheet, i, 0, header_text, line)
        else:
            wb.add_single_row_new_way(sheet, i, 0, plain_text, line)
        i += 1

    # Limit sheet
    sheet = wb.get_new_worksheet("Limits")
    update_labels_in_dictdata(limits, label_mapping)
    list1 = list(limits.items())
    wb.fill_sheet(sheet, col_dict_level_0, list1)

    # Create Parsed Sheets from Plugin Info
    for typename in type_names:
        row = 0
        lm = label_mapping.copy()
        lm = update_external_internal_in_label_mapping(typename, lm)
        sheet_name = (typename[0]+'-'+typename[1])[:31]
        sheet = wb.get_new_worksheet(sheet_name)
        attrdict = types[typename]['input']
        fieldmappingslist = attrdict['FieldMappings']
        attrdict.pop('FieldMappings')
        wb.add_headers(sheet, col_dict_level_0, 2)
        row = row + 1
        taskmappings = {}
        for key, value in attrdict.items():
            if 'Conditions' in key and len(value) != 0:
                row = wb.add_single_row(
                    sheet, row, col_dict_level_0, (update_label(key, lm), ':'))
                row = write_conditions(value, lm, row)
            elif type(value) is dict and len(value) == 0:
                row = wb.add_single_row(
                    sheet, row, col_dict_level_0, (update_label(key, lm), '-'))
            elif type(value) is dict and len(value) > 0:
                taskmappings = {key: value}
            elif type(value) is bool and value is True:
                row = wb.add_single_row(
                    sheet, row, col_dict_level_0, (update_label(key, lm), unicode_symbols["tick"]))
            elif type(value) is bool and value is False:
                row = wb.add_single_row(
                    sheet, row, col_dict_level_0, (update_label(key, lm), unicode_symbols["cross"]))
            else:
                row = wb.add_single_row(
                    sheet, row, col_dict_level_0, (update_label(key, lm), value))

        if len(taskmappings) > 0:
            res = list(taskmappings.keys())[0]
            row = wb.add_single_row(
                sheet, row, col_dict_level_0, (update_label(res, lm), ':'))
            for item in taskmappings[res]:
                value = taskmappings[res][item]
                if type(value) is bool and value is True:
                    row = wb.add_single_row(
                        sheet, row, col_dict_task_mapping, (update_label(item, lm),  unicode_symbols["tick"]))
                elif type(value) is bool and value is False:
                    row = wb.add_single_row(
                        sheet, row, col_dict_task_mapping, (update_label(item, lm),  unicode_symbols["cross"]))
                else:
                    row = wb.add_single_row(
                        sheet, row, col_dict_task_mapping, (update_label(item, lm), value))

        # external_name = (typename[1])[:31]
        fm_sheet_name = typename[0][0] + "-" + \
            typename[1][0:13] + " Field Mappings"
        sheet = wb.get_new_worksheet(fm_sheet_name)
        df_fm = pd.DataFrame(columns=list(field_mapping.keys()))
        df_fm = df_fm.append(fieldmappingslist, ignore_index=True)
        df_fm.fillna('', inplace=True)

        listoffieldmappings = df_fm.values.tolist()
        filtered_listoffieldmappings_list = []

        for i in listoffieldmappings:
            temp = []
            temp.append(i[0])  # Outreach Field Name 0
            temp.append(i[2])  # SF Field Name 1
            temp.append('')  # Left for Field Type...TODO: should be dropdown 2
            # Left for Record Type...TODO: should be dropdown 3
            temp.append('')
            # temp.append('') ## Recommended empty or prefilled 4
            # temp.append('')  ## UI Visibility TODO: pre set values 5
            if i[10] == True:  # Updates IN 6
                temp.append(unicode_symbols["tick"])  # UI
            else:
                temp.append('')
            if i[11] == True:  # Updates OUT 7
                temp.append(unicode_symbols["tick"])  # UI check
            else:
                temp.append('')
            temp.append('')  # NOTES 8
            temp.append('')  # reserved for popup notes 9
            temp.append('')  # reserved for
            temp.append('')
            temp.append('')
            temp.append('')
            temp.append('')
            filtered_listoffieldmappings_list.append(temp)
        if typename[0] in types_mapping_to_preset_data.keys():
            temp_preset = types_mapping_to_preset_data[typename[0]]
            for i in filtered_listoffieldmappings_list:
                if i[0] in temp_preset.keys():
                    index = filtered_listoffieldmappings_list.index(i)
                    filtered_listoffieldmappings_list[index][2] = temp_preset[i[0]]["FieldType"]
                    filtered_listoffieldmappings_list[index][3] = temp_preset[i[0]]["RecordType"]
                    # filtered_listoffieldmappings_list[index][4] = temp_preset[i[0]]["Recommended"]
                    # filtered_listoffieldmappings_list[index][5] = temp_preset[i[0]]["UI Visibility"]
                    filtered_listoffieldmappings_list[index][12] = temp_preset[i[0]]["Note"]

        wb.fill_sheet(sheet, col_field_mapping1,
                      filtered_listoffieldmappings_list)

    wb.close_workbook()

    autofit_spreadsheet_columns(spreadsheet_filename)
